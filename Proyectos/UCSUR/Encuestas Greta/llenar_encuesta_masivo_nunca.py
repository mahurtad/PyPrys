import asyncio
import pandas as pd
from datetime import datetime
import csv
import time
from playwright.async_api import async_playwright
from tqdm import tqdm

# ======= CONFIGURACIÓN =======
NUM_ENCUESTAS = 600  # Total de encuestas a enviar
NUM_HILOS = 10      # Usuarios simultáneos
FORM_URL = "https://forms.office.com/r/YQuDbeVfJz"

# Configurar log
log_path = "encuestas_log.csv"
with open(log_path, mode="w", newline="", encoding="utf-8") as logfile:
    writer = csv.DictWriter(logfile, fieldnames=["timestamp", "encuesta_num", "estado", "detalle"])
    writer.writeheader()

# ======= FUNCIÓN PRINCIPAL =======
async def enviar_encuesta(encuesta_num):
    estado = "OK"
    detalle = ""
    
    try:
        async with async_playwright() as p:
            # Configuración headless sin interfaz gráfica
            browser = await p.chromium.launch(
                headless=True,
                args=[
                    '--no-sandbox',
                    '--disable-setuid-sandbox',
                    '--disable-dev-shm-usage',
                    '--disable-accelerated-2d-canvas',
                    '--disable-gpu'
                ]
            )
            context = await browser.new_context(
                viewport={'width': 1920, 'height': 1080},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36'
            )
            page = await context.new_page()
            page.set_default_timeout(30000)
            page.set_default_navigation_timeout(30000)

            # Navegación inicial
            await page.goto(FORM_URL, wait_until="networkidle")
            await page.get_by_role("button", name="Start now").click()

            # === SECCIÓN DE PREGUNTAS ===

            # Pregunta 1 - Seleccionar primera opción
            await page.get_by_role("radio", name="Esta es la primera vez").click()
            
            # Pregunta 2 - Seleccionar "Nunca" (modificación solicitada)
            await page.get_by_role("radio", name="Nunca").click()

            # Pregunta 3 - Rellenar motivo (aparece al seleccionar "Nunca")
            motivo = "No necesité usar el tutor virtual"
            await page.get_by_label("En caso no lo hayas utilizado, ¿cuál es el motivo?").fill(motivo)

            # Enviar formulario inmediatamente (sin responder otras preguntas)
            await page.get_by_role("button", name="Submit").click()
            await page.wait_for_timeout(2000)  # Espera para confirmar envío

            await context.close()
            await browser.close()
            
            print(f"✅ Encuesta {encuesta_num} completada")
            
    except Exception as e:
        estado = "ERROR"
        detalle = str(e)
        print(f"❌ Encuesta {encuesta_num} fallida: {str(e)}")
    
    # Registrar en log
    with open(log_path, mode="a", newline="", encoding="utf-8") as logfile:
        writer = csv.DictWriter(logfile, fieldnames=["timestamp", "encuesta_num", "estado", "detalle"])
        writer.writerow({
            "timestamp": datetime.now().isoformat(),
            "encuesta_num": encuesta_num,
            "estado": estado,
            "detalle": detalle
        })
    return estado

# ======= EJECUCIÓN PRINCIPAL =======
async def main():
    print(f"🚀 Iniciando envío de {NUM_ENCUESTAS} encuestas (modo headless) con {NUM_HILOS} hilos simultáneos\n")
    inicio_tiempo = time.time()
    
    # Control de concurrencia
    semaphore = asyncio.Semaphore(NUM_HILOS)
    
    async def worker(encuesta_num):
        async with semaphore:
            return await enviar_encuesta(encuesta_num)
    
    # Ejecutar encuestas con barra de progreso
    resultados = []
    with tqdm(total=NUM_ENCUESTAS, desc="Progreso", unit="enc") as pbar:
        for i in range(0, NUM_ENCUESTAS, NUM_HILOS):
            batch = range(i+1, min(i+1+NUM_HILOS, NUM_ENCUESTAS+1))
            tasks = [worker(num) for num in batch]
            resultados.extend(await asyncio.gather(*tasks))
            pbar.update(len(batch))
    
    # Resumen final
    tiempo_total = (time.time() - inicio_tiempo) / 60
    print(f"\n✅ Proceso completado en {tiempo_total:.1f} minutos")
    print(f"📊 Resultados guardados en: {log_path}")
    
    # Estadísticas
    exitos = resultados.count("OK")
    errores = resultados.count("ERROR")
    print(f"\nResumen: {exitos} éxitos | {errores} errores")
    
    if errores > 0:
        print("\n🔍 Detalle de errores:")
        df_log = pd.read_csv(log_path)
        print(df_log[df_log["estado"] == "ERROR"][["encuesta_num", "detalle"]].to_string(index=False))

if __name__ == "__main__":
    asyncio.run(main())