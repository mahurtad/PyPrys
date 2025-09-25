# -*- coding: utf-8 -*-
"""
Canvas (SAML Microsoft) -> New Analytics (LTI 212) en la MISMA pestaña.
Cierra popups pulsando 'Space' (2-4 veces), activa 'Actividad semanal en línea',
extrae la fila (p.ej. "Video Asíncrono 1") y exporta a Excel.

Requiere: playwright, pandas, openpyxl
pip install playwright pandas openpyxl
python -m playwright install
"""

import os
import re
import asyncio
import pandas as pd
from typing import Optional, Tuple
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# =============== Config ===============
BASE_URL = "https://cientificavirtual.cientifica.edu.pe"
CANVAS_LOGIN_URL = f"{BASE_URL}/login/canvas"
ANALYTICS_TOOL_ID = "212"  # New Analytics (LTI)

COURSE_ID = int(os.getenv("COURSE_ID", "58735"))
RESOURCE_NAME = os.getenv("RESOURCE_NAME", "Video Asíncrono 1").strip()
OUT_XLSX = f"reporte_analiticas_{COURSE_ID}.xlsx"

USER_XLSX = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\automatix\proyectos_api_canvas_openapi\user.xlsx"
# USER_XLSX = "G:/My Drive/Data Analysis/Proyectos/UCSUR/Rubricas_UCSUR/automatix/proyectos_api_canvas_openapi/user.xlsx"

# =============== Utilidades ===============
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

# =============== Login SAML (Microsoft) ===============
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

    await page.wait_for_url(re.compile(r"cientificavirtual\.cientifica\.edu\.pe", re.I), timeout=60000)
    return page

# =============== Navegar a Analytics (misma pestaña) ===============
async def goto_analytics_same_page(page, course_id: int):
    url = f"{BASE_URL}/courses/{course_id}/external_tools/{ANALYTICS_TOOL_ID}?launch_type=course_navigation"
    try:
        await page.goto(url, wait_until="domcontentloaded", timeout=120_000)
    except PWTimeout:
        pass
    await page.wait_for_selector("text=Analíticas del curso", timeout=120_000)

# =============== Cerrar popups pulsando SPACE ===============
async def close_popups_with_space(target, presses: int = 2):
    """
    Algunas instalaciones de Canvas/Reactour aceptan 'Space' como acción por defecto del diálogo.
    Forzamos foco al body o al botón primario y mandamos SPACE varias veces.
    """
    # intenta enfocar botón por si existe
    try:
        btn = target.locator(
            "button:has-text('Ahora no'), [role='button']:has-text('Ahora no'), "
            "div.reactour__helper[role='dialog'] button, "
            "div[role='dialog'] button"
        ).first
        if await btn.count():
            try: await btn.focus()
            except Exception: pass
    except Exception:
        pass

    # asegurar foco en body para que reciba la tecla
    try:
        await target.evaluate("""
        (() => {
          if (!document.body.hasAttribute('tabindex')) document.body.setAttribute('tabindex','-1');
          document.body.focus();
        })();
        """)
    except Exception:
        pass

    # mandar SPACE (y como refuerzo, Escape entre medias)
    for i in range(max(2, presses)):
        try: await target.keyboard.press("Space")
        except Exception: pass
        await target.wait_for_timeout(200)
        try: await target.keyboard.press("Escape")
        except Exception: pass
        await target.wait_for_timeout(150)

    # por si hay iframes con el foco dentro
    try:
        for fr in target.frames:
            await close_popups_with_space(fr, presses=1)
    except Exception:
        pass

# =============== Activar pestaña Actividad semanal ===============
async def activate_online_activity_tab(page):
    await close_popups_with_space(page, presses=3)

    async def activate_in(target) -> bool:
        if await target.locator("#tab-participation").count():
            tab = target.locator("#tab-participation").first
        elif await target.locator("[role='tab'][aria-controls='participation']").count():
            tab = target.locator("[role='tab'][aria-controls='participation']").first
        else:
            tab = target.get_by_role("tab", name=re.compile(r"Actividad semanal en l[íi]nea", re.I)).first
            if not await tab.count(): return False

        # click normal → force
        try:
            await tab.scroll_into_view_if_needed(); await tab.click(timeout=1200)
        except Exception:
            try: await tab.click(timeout=1200, force=True)
            except Exception: pass
        try:
            if (await tab.get_attribute("aria-selected")) == "true": return True
        except Exception: pass

        # teclado sobre el tablist
        try:
            tl = target.locator("[role='tablist']").first
            if await tl.count():
                await tl.focus()
                for _ in range(6):
                    await target.keyboard.press("ArrowRight")
                    await target.wait_for_timeout(90)
                    if await target.locator("#tab-participation[aria-selected='true']").count():
                        return True
        except Exception: pass

        # JS click
        try:
            handle = await tab.element_handle()
            if handle:
                await target.evaluate("(el)=>{ el.click(); el.setAttribute('aria-selected','true'); }", handle)
                if await target.locator("#tab-participation[aria-selected='true']").count():
                    return True
        except Exception: pass

        return False

    ok = await activate_in(page)
    if not ok:
        # pulsar SPACE otra vez por si el foco se quedó en el modal
        await close_popups_with_space(page, presses=2)
        ok = await activate_in(page)

    if not ok:
        try:
            await page.reload(wait_until="domcontentloaded", timeout=120_000)
        except PWTimeout:
            pass
        await page.wait_for_selector("text=Analíticas del curso", timeout=120_000)
        await close_popups_with_space(page, presses=3)
        if not await activate_in(page):
            raise RuntimeError("No se pudo activar la pestaña 'Actividad semanal en línea'.")

    await page.wait_for_selector("table thead >> text=Recurso", timeout=60_000)

# =============== Extraer fila del recurso ===============
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
        i_name = idx("recurso"); i_stu = idx("estudiantes")
        i_view = idx("vistas");  i_part = idx("participaciones")
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

# =============== Guardar Excel ===============
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

# =============== Main ===============
async def main():
    email, password = read_credentials_from_excel(USER_XLSX)

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(viewport={"width":1366, "height":768})
        context.set_default_navigation_timeout(120_000)
        context.set_default_timeout(60_000)

        # (opcional) bloquear overlays de Qualtrics si aparecieran
        await context.route("**/*siteintercept.qualtrics.com/**", lambda route: asyncio.create_task(route.abort()))

        page = await login_saml_into_canvas(context, email, password)
        await goto_analytics_same_page(page, COURSE_ID)

        # cerrar popups con SPACE (2 veces suele bastar)
        await close_popups_with_space(page, presses=3)

        # activar pestaña
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
