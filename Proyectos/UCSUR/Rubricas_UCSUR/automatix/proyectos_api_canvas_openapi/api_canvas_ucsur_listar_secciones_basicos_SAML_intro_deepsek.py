# -*- coding: utf-8 -*-
r"""
Canvas (SAML Microsoft) -> New Analytics (LTI 212) en la MISMA pestaña
Activa la pestaña 'Actividad semanal en línea', extrae la fila del recurso
(ej. "Video Asíncrono 1") y exporta a Excel.

Credenciales en Excel (columnas: 'email', 'contraseña' o 'contrasena'):
G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\automatix\proyectos_api_canvas_openapi\user.xlsx
"""

import os
import re
import asyncio
import pandas as pd
from typing import Optional, Tuple
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============================ Config ============================

BASE_URL = "https://cientificavirtual.cientifica.edu.pe"
CANVAS_LOGIN_URL = f"{BASE_URL}/login/canvas"
ANALYTICS_TOOL_ID = "212"  # New Analytics (LTI)

COURSE_ID = int(os.getenv("COURSE_ID", "58735"))
RESOURCE_NAME = os.getenv("RESOURCE_NAME", "Video Asíncrono 1").strip()
OUT_XLSX = f"reporte_analiticas_{COURSE_ID}.xlsx"

# Excel con credenciales (usa raw string o / para evitar \U escapes)
USER_XLSX = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\automatix\proyectos_api_canvas_openapi\user.xlsx"
# USER_XLSX = "G:/My Drive/Data Analysis/Proyectos/UCSUR/Rubricas_UCSUR/automatix/proyectos_api_canvas_openapi/user.xlsx"

# ======================= Utilidades Excel =======================

def _norm(s: str) -> str:
    s = s.strip().lower()
    for a,b in (("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")): s = s.replace(a,b)
    return s

def read_credentials_from_excel(path: str) -> Tuple[str, str]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"No existe el archivo de credenciales: {path}")
    df = pd.read_excel(path, engine="openpyxl")
    df.rename(columns={c: _norm(str(c)) for c in df.columns}, inplace=True)
    pwd_col = "contraseña" if "contraseña" in df.columns else ("contrasena" if "contrasena" in df.columns else None)
    if "email" not in df.columns or not pwd_col:
        raise ValueError("El Excel debe tener columnas 'email' y 'contraseña' (o 'contrasena').")
    row = df.loc[(df["email"].astype(str).str.len()>0) & (df[pwd_col].astype(str).str.len()>0)].iloc[0]
    return str(row["email"]).strip(), str(row[pwd_col]).strip()

# ======================= Login SAML (MS) ========================

async def _set_input(page, selector: str, value: str, timeout=12000) -> bool:
    el = page.locator(selector).first
    await el.wait_for(state="visible", timeout=timeout)
    try: await el.click(timeout=timeout)
    except Exception: pass
    try:
        await el.fill("", timeout=timeout); await el.type(value, delay=25)
    except Exception:
        try: await el.fill(value, timeout=timeout)
        except Exception: pass
    try: return (await el.input_value()) == value
    except Exception: return False

async def _email_step(page, email: str) -> bool:
    field = "input[name='loginfmt'], #i0116"
    btn   = "#idSIButton9, button:has-text('Siguiente'), button:has-text('Next')"
    for _ in range(3):
        if not await _set_input(page, field, email): continue
        try: await page.click(btn, timeout=8000)
        except Exception: pass
        try:
            await page.wait_for_selector("input[name='passwd'], #i0118", state="visible", timeout=10000)
            return True
        except PWTimeout:
            if "login.microsoftonline.com" not in (page.url or ""):
                return True
    return False

async def _password_step(page, password: str) -> bool:
    field = "input[name='passwd'], #i0118"
    btn   = "#idSIButton9, button:has-text('Iniciar sesión'), button:has-text('Sign in')"
    for _ in range(3):
        if not await _set_input(page, field, password): continue
        try: await page.click(btn, timeout=8000)
        except Exception: pass
        try:
            await page.wait_for_url(re.compile(r"microsoftonline\.com|cientifica\.edu\.pe", re.I), timeout=12000)
            return True
        except PWTimeout:
            return True

async def _kmsi_no(page):
    try:
        loc = page.locator("#idBtn_Back, button:has-text('No')")
        if await loc.first.is_visible(timeout=3000):
            await loc.first.click(timeout=3000)
    except Exception: pass

async def login_saml_into_canvas(context, email: str, password: str):
    page = await context.new_page()
    await page.goto(CANVAS_LOGIN_URL, wait_until="domcontentloaded")
    try:
        await page.get_by_role("link", name=re.compile(r"Iniciar sesión.*Microsoft", re.I)).click()
    except PWTimeout:
        await page.locator("text=Iniciar sesión con Microsoft").first.click()

    await page.wait_for_url(re.compile(r"login\.microsoftonline\.com", re.I), timeout=30000)
    if not await _email_step(page, email):      raise RuntimeError("Fallo en email_step.")
    if not await _password_step(page, password): raise RuntimeError("Fallo en password_step.")
    await _kmsi_no(page)

    # nos quedamos en la pestaña actual; aseguramos que ya volvimos a Canvas
    await page.wait_for_url(re.compile(r"cientificavirtual\.cientifica\.edu\.pe", re.I), timeout=60000)
    return page

# ============== Navegar a Analytics en la MISMA pestaña =========

async def goto_analytics_same_page(page, course_id: int):
    url = f"{BASE_URL}/courses/{course_id}/external_tools/{ANALYTICS_TOOL_ID}?launch_type=course_navigation"
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    except PWTimeout:
        pass
    await page.wait_for_selector("text=Analíticas del curso", timeout=120_000)

# =========== Neutralizar/eliminar popups & overlays ============

async def dismiss_welcome_popup(target):
    """Función mejorada para manejar popups de Canvas con teclado"""
    print("🔍 Buscando y cerrando popups...")
    
    # PRIMERO: Intentar con la barra espaciadora (como mencionaste que funciona)
    print("⌨️  Intentando cerrar popups con barra espaciadora...")
    for _ in range(3):  # Intentar múltiples veces
        try:
            # Enfocar el body primero para asegurar que los eventos de teclado lleguen
            await target.evaluate("() => document.body.focus()")
            await target.keyboard.press("Space")
            print("✅ Barra espaciadora presionada")
            await asyncio.sleep(0.5)  # Pequeña pausa entre pulsaciones
            
            # También probar con Escape que es común para cerrar modales
            await target.keyboard.press("Escape")
            print("✅ Tecla Escape presionada")
            await asyncio.sleep(0.5)
        except Exception as e:
            print(f"❌ Error al presionar teclas: {e}")
    
    # SEGUNDO: Buscar botones específicos de cierre
    close_selectors = [
        "button:has-text('Ahora no')",
        "button:has-text('No ahora')", 
        "button:has-text('Later')",
        "button:has-text('Not now')",
        "button:has-text('Omitir')",
        "button:has-text('Skip')",
        "button:has-text('Cerrar')",
        "button:has-text('Close')",
        "button[aria-label*='Cerrar']",
        "button[aria-label*='Close']",
        "button:has-text('×')",
        ".close-button",
        ".modal-close",
        ".ReactModal__Close"
    ]
    
    for selector in close_selectors:
        try:
            buttons = target.locator(selector)
            count = await buttons.count()
            if count > 0:
                print(f"✅ Encontrados {count} botones con selector: {selector}")
                # Intentar hacer clic en todos los botones encontrados
                for i in range(count):
                    try:
                        btn = buttons.nth(i)
                        if await btn.is_visible(timeout=1000):
                            await btn.click(timeout=1000)
                            print(f"✅ Botón {i+1} clickeado")
                            await asyncio.sleep(0.3)
                    except Exception:
                        try:
                            btn = buttons.nth(i)
                            if await btn.is_visible(timeout=1000):
                                await btn.click(timeout=1000, force=True)
                                print(f"✅ Botón {i+1} clickeado (force)")
                                await asyncio.sleep(0.3)
                        except Exception:
                            continue
                return True
        except Exception:
            continue
    
    # TERCERO: Buscar por texto en popups
    popup_texts = [
        "Recorrido del administrador",
        "Welcome to Canvas",
        "Bienvenido a Canvas", 
        "Comienza el recorrido",
        "Tour de administración",
        "Ahora no",
        "Comenzar recorrido"
    ]
    
    for text in popup_texts:
        try:
            elements = target.get_by_text(text, exact=False)
            count = await elements.count()
            if count > 0:
                print(f"✅ Encontrados {count} elementos con texto: {text}")
                # Buscar botones de cierre cerca de estos textos
                for i in range(count):
                    try:
                        element = elements.nth(i)
                        if await element.is_visible(timeout=1000):
                            # Buscar botones cercanos
                            close_btn = element.locator("xpath=./ancestor::*[contains(@class, 'modal') or contains(@class, 'popup')]//button[contains(text(), 'Ahora no') or contains(text(), 'Cerrar') or contains(text(), 'Close')]")
                            if await close_btn.count() > 0:
                                await close_btn.first.click(timeout=1000)
                                print("✅ Popup cerrado por texto")
                                return True
                    except Exception:
                        continue
        except Exception:
            continue
    
    # CUARTO: Eliminación nuclear como último recurso
    try:
        removed_count = await target.evaluate("""
        () => {
            const selectors = [
                '[role="dialog"]', '.ReactModal__Overlay', '.ReactModalPortal',
                '.reactour__helper', '#reactour', '.shepherd-modal-overlay-container',
                '.coach-mark', '.tour', '.ui-popover', '.modal-backdrop',
                '.ql-overlay', '#QSISSurveyWindowPopup', '.popup-overlay',
                '.onboarding-modal', '.introjs-overlay', '.joyride-overlay'
            ];
            
            let count = 0;
            selectors.forEach(selector => {
                document.querySelectorAll(selector).forEach(el => {
                    try {
                        el.remove();
                        count++;
                    } catch(e) {}
                });
            });
            
            // Restaurar funcionalidad de la página
            document.documentElement.style.overflow = 'auto';
            document.body.style.overflow = 'auto';
            document.body.style.pointerEvents = 'auto';
            
            return count;
        }
        """)
        
        if removed_count > 0:
            print(f"✅ Eliminados {removed_count} elementos overlay")
            return True
    except Exception as e:
        print(f"❌ Error en eliminación nuclear: {e}")
    
    print("⚠️  No se encontraron popups visibles o no se pudieron cerrar")
    return False

# ======== Activar pestaña "Actividad semanal en línea" =========

async def activate_online_activity_tab(page):
    MAX_RETRIES = 3
    
    for attempt in range(MAX_RETRIES):
        print(f"🔄 Intento {attempt + 1} de {MAX_RETRIES} para activar pestaña")
        
        # Cerrar popups antes de cada intento
        for _ in range(2):
            await dismiss_welcome_popup(page)
            await asyncio.sleep(1)
        
        # Estrategia múltiple para encontrar la pestaña
        tab_selectors = [
            # Selectores principales
            "#tab-participation",
            "[role='tab'][aria-controls='participation']",
            "button[data-tab='participation']",
            "a[href*='participation']",
            
            # Selectores por texto
            "//*[contains(text(), 'Actividad semanal en línea') or contains(text(), 'Online Weekly Activity')]",
            "button:has-text('Actividad semanal en línea'), button:has-text('Online Weekly Activity')",
            "a:has-text('Actividad semanal en línea'), a:has-text('Online Weekly Activity')",
            
            # Selectores por atributos ARIA
            "[aria-label*='Actividad semanal' i], [aria-label*='Weekly Activity' i]"
        ]
        
        tab_found = None
        for selector in tab_selectors:
            try:
                if selector.startswith("//"):
                    tab = page.locator("xpath=" + selector)
                else:
                    tab = page.locator(selector).first
                
                if await tab.count() > 0 and await tab.is_visible(timeout=3000):
                    tab_found = tab
                    print(f"✅ Pestaña encontrada con selector: {selector}")
                    break
            except Exception:
                continue
        
        # Buscar en iframes si no se encuentra en página principal
        if not tab_found:
            print("🔍 Buscando pestaña en iframes...")
            for frame in page.frames:
                if frame != page.main_frame:
                    try:
                        for selector in tab_selectors:
                            if selector.startswith("//"):
                                tab = frame.locator("xpath=" + selector)
                            else:
                                tab = frame.locator(selector).first
                            
                            if await tab.count() > 0 and await tab.is_visible(timeout=2000):
                                tab_found = tab
                                print(f"✅ Pestaña encontrada en iframe con selector: {selector}")
                                break
                        if tab_found:
                            break
                    except Exception:
                        continue
        
        if not tab_found:
            print("❌ No se pudo encontrar la pestaña")
            if attempt < MAX_RETRIES - 1:
                await page.reload(wait_until="domcontentloaded", timeout=120000)
                await page.wait_for_selector("text=Analíticas del curso", timeout=60000)
                continue
            else:
                raise RuntimeError("No se pudo encontrar la pestaña 'Actividad semanal en línea'.")
        
        # Intentar activar la pestaña
        try:
            await tab_found.scroll_into_view_if_needed()
            await tab_found.click(timeout=5000)
            print("✅ Clic en pestaña realizado")
            
            await asyncio.sleep(2)
            
            # Verificar activación
            try:
                await page.wait_for_selector("table thead >> text=Recurso", timeout=10000)
                print("✅ Pestaña activada correctamente")
                return
            except Exception:
                # Intentar con JavaScript si el click normal no funciona
                await page.evaluate("""(selector) => {
                    let element;
                    if (selector.startsWith('//')) {
                        element = document.evaluate(selector, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    } else {
                        element = document.querySelector(selector);
                    }
                    if (element) {
                        element.click();
                        return true;
                    }
                    return false;
                }""", selector)
                print("✅ Pestaña activada mediante JavaScript")
                await asyncio.sleep(2)
                return
                
        except Exception as e:
            print(f"❌ Error al activar pestaña: {e}")
            if attempt < MAX_RETRIES - 1:
                await page.reload(wait_until="domcontentloaded", timeout=120000)
                await page.wait_for_selector("text=Analíticas del curso", timeout=60000)
            else:
                raise RuntimeError("No se pudo activar la pestaña 'Actividad semanal en línea'.")

    # Confirmación final
    try:
        await page.wait_for_selector("table thead >> text=Recurso", timeout=30000)
        print("✅ Tabla de recursos cargada correctamente")
    except Exception:
        print("⚠️  No se pudo confirmar la carga de la tabla")

# =================== Leer fila de recurso =======================

async def read_resource_row(page, resource_name: str) -> Optional[Tuple[str, str, str, str]]:
    try:
        course_title = await page.locator("h1:has-text('Analíticas del curso')").first.inner_text()
    except Exception:
        course_title = "Curso"

    async def extract_current_page():
        headers = [await h.inner_text() for h in page.locator("table thead th").all()]
        headers = [h.strip() for h in headers]
        def idx(label):
            for i,h in enumerate(headers):
                if label.lower() in h.lower(): return i
            return None
        i_name = idx("Recurso"); i_stu = idx("Estudiantes")
        i_view = idx("Vistas de página"); i_part = idx("Participaciones")
        if None in (i_name, i_stu, i_view, i_part): return None
        rows = page.locator("table tbody tr"); count = await rows.count()
        for i in range(count):
            tds = rows.nth(i).locator("td")
            name = (await tds.nth(i_name).inner_text()).strip()
            if name == resource_name:
                estudiantes   = (await tds.nth(i_stu).inner_text()).strip()
                vistas        = (await tds.nth(i_view).inner_text()).strip()
                participacion = (await tds.nth(i_part).inner_text()).strip()
                return course_title, estudiantes, vistas, participacion
        return None

    for _ in range(11):
        got = await extract_current_page()
        if got: return got
        next_btn = page.get_by_role("link", name=re.compile(r"(›|Siguiente|Página siguiente)", re.I))
        if await next_btn.count() > 0 and await next_btn.first.is_enabled():
            await next_btn.first.click(); await page.wait_for_timeout(650)
        else: break
    return None

# ====================== Exportar a Excel ========================

def save_excel(course_id: int, course_name: str, resource: str,
               estudiantes: str, vistas: str, particip: str, path: str):
    df = pd.DataFrame([{
        "course_id": course_id,
        "course_name": course_name.replace("Analíticas del curso", "").strip() or course_name.strip(),
        "Nombre de recurso": resource,
        "Estudiantes": estudiantes,
        "Vistas de página": vistas,
        "Participaciones": particip
    }])
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Reporte")
        ws = xw.sheets["Reporte"]
        for c_idx, col in enumerate(df.columns, start=1):
            max_len = max(len(str(col)), *(len(str(v)) for v in df[col].astype(str)))
            ws.column_dimensions[ws.cell(row=1, column=c_idx).column_letter].width = min(max_len + 2, 60)

# ============================ Main =============================

async def main():
    email, password = read_credentials_from_excel(USER_XLSX)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width":1366, "height":768})
        context.set_default_navigation_timeout(120_000)
        context.set_default_timeout(60_000)

        page = await login_saml_into_canvas(context, email, password)
        await goto_analytics_same_page(page, COURSE_ID)

        # Cerrar popups de manera agresiva (múltiples intentos)
        print("🚀 Cerrando popups de manera agresiva...")
        for i in range(5):  # Hasta 5 intentos
            print(f"🔄 Intento {i+1} de 5 para cerrar popups")
            closed = await dismiss_welcome_popup(page)
            if closed:
                print(f"✅ Popups cerrados en intento {i+1}")
            await asyncio.sleep(1)

        # Activar pestaña
        await activate_online_activity_tab(page)

        row = await read_resource_row(page, RESOURCE_NAME)
        if not row:
            print(f"❗ No se encontró la fila '{RESOURCE_NAME}' en el curso {COURSE_ID}.")
            await browser.close(); return
        course_title, estudiantes, vistas, particip = row

        save_excel(COURSE_ID, course_title, RESOURCE_NAME, estudiantes, vistas, particip, OUT_XLSX)
        print(f"✅ Listo: {OUT_XLSX}")

        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())