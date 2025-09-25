import asyncio
import random
import pandas as pd
from datetime import datetime
import csv
import time
from playwright.async_api import async_playwright
from tqdm import tqdm

# ======= CONFIGURACIÓN =======
NUM_ENCUESTAS = 50  # Total de encuestas a enviar
NUM_HILOS = 5       # Usuarios simultáneos (3-5 recomendado)
FORM_URL = "https://forms.office.com/r/YQuDbeVfJz"

# Archivos de respuestas (ajustar rutas según necesidad)
RESPUESTAS_19 = "pregunta_19.xlsx"
RESPUESTAS_21 = "pregunta_21.xlsx"
RESPUESTAS_22 = "pregunta_22.xlsx"

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
            # Configuración headless optimizada
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

            # Pregunta 1
            await page.get_by_role("radio", name="Esta es la primera vez").click()
            
            # Pregunta 2
            opciones_2 = [
                "Muy frecuentemente (Más de 10 veces en total)",
                "Frecuentemente (6-10 veces en total)",
                "Ocasionalmente (3-5 veces en total)",
                "Rara vez (1-2 veces en total)"
            ]
            await page.get_by_role("radio", name=random.choice(opciones_2)).click()

            # Pregunta 3
            opciones_3 = [
                "Menos de 5 minutos",
                "Entre 5 y 10 minutos",
                "Entre 10 y 20 minutos",
                "Más de 20 minutos"
            ]
            await page.get_by_role("radio", name=random.choice(opciones_3)).click()
            
            # Pregunta 4
            await page.get_by_role("radio", name="Solo para Química").click()
            
            # Pregunta 5 (selección múltiple)
            opciones_5 = [
                "Obtener resúmenes de los contenidos del curso",
                "Obtener ejercicios o cuestionarios para practicar",
                "Obtener explicaciones sobre conceptos o temas del curso",
                "Obtener videos de los temas del curso",
                "Obtener repasos para mis evaluaciones",
                "Obtener ejemplos prácticos sobre conceptos o temas del curso",
                "Obtener recordatorios sobre las fechas de mis evaluaciones"
            ]
            for opcion in random.sample(opciones_5, random.randint(1, 4)):
                await page.get_by_role("checkbox", name=opcion).click()
            
            # Pregunta 6
            await page.get_by_label("¿Tu docente te ha comentado sobre la existencia del tutor virtual y ha promovido su uso en clase?")\
                .locator("..").get_by_text("Sí", exact=True).click()
            
            # Pregunta 7
            opcion_utilidad = random.choice(["Muy útiles", "Algo útiles"])
            await page.get_by_label("¿Qué tan útiles encuentras las explicaciones del tutor virtual para comprender conceptos complejos de Química?")\
                .locator("..").get_by_text(opcion_utilidad, exact=True).click()
            
            # Pregunta 8
            opcion_conformidad = random.choice(["Muy conforme", "Conforme"])
            await page.get_by_label("¿Estás conforme con la calidad de los ejercicios e información que el tutor virtual ha ofrecido?")\
                .locator("..").get_by_text(opcion_conformidad, exact=True).click()

            # Pregunta 9
            opcion_p9 = random.choice(["Muy conforme", "Conforme"])
            await page.get_by_label("¿Estás conforme con la calidad de los ejemplos prácticos que el tutor virtual ofrece para reforzar los temas?")\
                .locator("..").get_by_text(opcion_p9, exact=True).click()

            # Pregunta 10
            opcion_p10 = random.choice(["Muy útiles", "Algo útiles"])
            await page.get_by_label("¿Qué tan útiles consideras los exámenes/ cuestionarios ofrecidos por el tutor virtual para prepararte?")\
                .locator("..").get_by_text(opcion_p10, exact=True).click()

            # Pregunta 11
            opcion_p11 = random.choice(["Muy útil", "Algo útil"])
            await page.get_by_label("¿Qué tan útil te resulta la retroalimentación que el tutor virtual ofrece cuando aciertas o fallas una pregunta?")\
                .locator("..").get_by_text(opcion_p11, exact=True).click()
            
            # Pregunta 12 (selección múltiple)
            opciones_12 = [
                "Busqué la respuesta en mis apuntes o libros.",
                "Consulté con un profesor.",
                "Pedí ayuda a un compañero.",
                "Usé otra herramienta en línea."
            ]
            for opcion in random.sample(opciones_12, k=random.randint(1, 3)):
                await page.get_by_role("checkbox", name=opcion).check()

            # Pregunta 13
            opciones_13 = ["Ha mejorado mucho", "Ha mejorado bastante", "Ha mejorado moderadamente"]
            await page.get_by_label("¿Sientes que con el uso del Tutor Virtual ha mejorado tu comprensión general de los temas de Química?")\
                .locator("..").get_by_text(random.choice(opciones_13), exact=True).click()

            # Pregunta 14
            opciones_14 = ["Totalmente", "Bastante", "Moderadamente"]
            await page.get_by_label("¿Crees que el tutor virtual ha complementado de manera efectiva las explicaciones dadas en clase?")\
                .locator("..").get_by_text(random.choice(opciones_14), exact=True).click()

            # Pregunta 15 (valoración por estrellas)
            estrella_15 = random.choice([3, 4, 5])
            await page.get_by_role("radiogroup", name="15. En una escala del 1 al 5")\
                .locator(f'[aria-label="{estrella_15} Star"]').click()

            # Pregunta 16 (valoración por estrellas)
            estrella_16 = random.choice([3, 4, 5])
            await page.get_by_role("radiogroup", name="16. En una escala del 1 al 5")\
                .locator(f'[aria-label="{estrella_16} Star"]').click()
            
            # Pregunta 17
            opciones_17 = ["Muy satisfecho", "Satisfecho"]
            await page.get_by_label("¿Qué tan satisfecho estas con el tutor virtual?")\
                .locator("..").get_by_text(random.choice(opciones_17), exact=True).click()
            
            # Pregunta 18
            opciones_18 = ["No, todo ha sido fácil de usar", "Sí, algunas dificultades menores"]
            await page.get_by_label("¿Has encontrado alguna dificultad para usar el tutor virtual en estas ocho semanas?")\
                .locator("..").get_by_text(random.choice(opciones_18), exact=True).click()

            # Pregunta 19 (opcional - texto libre)
            if random.random() >= 0.3:  # 70% de probabilidad de responder
                try:
                    df_19 = pd.read_excel(RESPUESTAS_19)
                    respuesta_19 = random.choice(df_19.iloc[:, 0].dropna().tolist())
                    await page.get_by_label("¿Qué tipo de dificultades haz encontrado?").fill(respuesta_19)
                except Exception as e:
                    detalle += f"| Error P19: {str(e)} "
            
            # Pregunta 20
            await page.get_by_label("¿Recomendarías el uso del Tutor Virtual a otros estudiantes del curso?")\
                .locator("..").get_by_text("Sí", exact=True).click()

            # Pregunta 21 (opcional - texto libre)
            if random.random() >= 0.5:  # 50% de probabilidad de responder
                try:
                    df_21 = pd.read_excel(RESPUESTAS_21)
                    respuesta_21 = random.choice(df_21.iloc[:, 0].dropna().tolist())
                    await page.get_by_placeholder("Enter your answer").nth(1).fill(respuesta_21)
                except Exception as e:
                    detalle += f"| Error P21: {str(e)} "

            # Pregunta 22 (opcional - texto libre)
            if random.random() >= 0.3:  # 70% de probabilidad de responder
                try:
                    df_22 = pd.read_excel(RESPUESTAS_22)
                    respuesta_22 = random.choice(df_22.iloc[:, 0].dropna().tolist())
                    await page.get_by_label("¿Qué características adicionales te gustaría que tuviera el Tutor Virtual?").fill(respuesta_22)
                except Exception as e:
                    detalle += f"| Error P22: {str(e)} "

            # Pregunta 23 (opcional - texto libre)
            if random.random() >= 0.3:  # 70% de probabilidad de responder
                cursos_23 = [
                    "Bioquímica", "Biología", "Física", "Matemática general",
                    "Estadística general", "Álgebra", "Matemática I",
                    "Matemática II", "Matemática III"
                ]
                respuesta_23 = "Sí, para " + random.choice(cursos_23)
                await page.get_by_label("¿Consideras que sería útil utilizar esta herramienta en otros cursos? Si es así, ¿en cuáles?").fill(respuesta_23)

            # Envío final
            await page.get_by_role("button", name="Submit").click()
            await page.wait_for_timeout(2000)  # Espera para confirmar envío

            await context.close()
            await browser.close()
            
            print(f"✅ Encuesta {encuesta_num} completada")
            
    except Exception as e:
        estado = "ERROR"
        detalle = str(e) + detalle
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
    
    # Procesamiento por lotes con barra de progreso
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