import asyncio
import random
import pandas as pd
from playwright.async_api import async_playwright

async def run():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()

        # Abrir formulario
        await page.goto("https://forms.office.com/r/YQuDbeVfJz")
        print("Abriendo el formulario...")

        # Clic en "Start now"
        await page.get_by_role("button", name="Start now").click()
        print("Clic en 'Start now' exitoso.")

        # Pregunta 1
        print("Respondiendo Pregunta 1...")
        await page.get_by_role("radio", name="Esta es la primera vez").click()
        await asyncio.sleep(0.5)

        # Pregunta 2
        print("Respondiendo Pregunta 2...")
        opciones_2 = [
            "Muy frecuentemente (Más de 10 veces en total)",
            "Frecuentemente (6-10 veces en total)",
            "Ocasionalmente (3-5 veces en total)",
            "Rara vez (1-2 veces en total)"
        ]
        opcion_2 = random.choice(opciones_2)
        await page.get_by_role("radio", name=opcion_2).click()
        await asyncio.sleep(0.5)

        # Pregunta 3
        print("Respondiendo Pregunta 3...")
        opciones_3 = [
            "Menos de 5 minutos",
            "Entre 5 y 10 minutos",
            "Entre 10 y 20 minutos",
            "Más de 20 minutos"
        ]
        opcion_3 = random.choice(opciones_3)
        await page.get_by_role("radio", name=opcion_3).click()
        await asyncio.sleep(5)
        
        # Pregunta 4
        print("Respondiendo Pregunta 4...")
        await page.get_by_role("radio", name="Solo para Química").click()
        await asyncio.sleep(0.5)
        
        # Pregunta 5
        print("Respondiendo Pregunta 5...")
        opciones_5 = [
            "Obtener resúmenes de los contenidos del curso",
            "Obtener ejercicios o cuestionarios para practicar",
            "Obtener explicaciones sobre conceptos o temas del curso",
            "Obtener videos de los temas del curso",
            "Obtener repasos para mis evaluaciones",
            "Obtener ejemplos prácticos sobre conceptos o temas del curso",
            "Obtener recordatorios sobre las fechas de mis evaluaciones"
        ]

        # Elegir de 1 a 4 opciones aleatorias
        seleccionadas_5 = random.sample(opciones_5, random.randint(1, 4))
        for opcion in seleccionadas_5:
            await page.get_by_role("checkbox", name=opcion).click()
            await asyncio.sleep(0.2)
            
        # Pregunta 6
        print("Respondiendo Pregunta 6...")
        
        await page.get_by_label("¿Tu docente te ha comentado sobre la existencia del tutor virtual y ha promovido su uso en clase?").locator("..").get_by_text("Sí", exact=True).click()
        await asyncio.sleep(0.5)
        

        # Pregunta 7
        print("Respondiendo Pregunta 7...")
        opcion_utilidad = random.choice(["Muy útiles", "Algo útiles"])
        await page.get_by_label("¿Qué tan útiles encuentras las explicaciones del tutor virtual para comprender conceptos complejos de Química?")\
          .locator("..").get_by_text(opcion_utilidad, exact=True).click()
          
          
        # Pregunta 8        

        opcion_conformidad = random.choice(["Muy conforme", "Conforme"])
        await page.get_by_label("¿Estás conforme con la calidad de los ejercicios e información que el tutor virtual ha ofrecido?")\
          .locator("..").get_by_text(opcion_conformidad, exact=True).click()

        # Pregunta 9
        print("Respondiendo Pregunta 9...")
        opcion_p9 = random.choice(["Muy conforme", "Conforme"])
        await page.get_by_label("¿Estás conforme con la calidad de los ejemplos prácticos que el tutor virtual ofrece para reforzar los temas?")\
            .locator("..").get_by_text(opcion_p9, exact=True).click()

        # Pregunta 10
        print("Respondiendo Pregunta 10...")
        opcion_p10 = random.choice(["Muy útiles", "Algo útiles"])
        await page.get_by_label("¿Qué tan útiles consideras los exámenes/ cuestionarios ofrecidos por el tutor virtual para prepararte?")\
            .locator("..").get_by_text(opcion_p10, exact=True).click()

        # Pregunta 11
        print("Respondiendo Pregunta 11...")
        opcion_p11 = random.choice(["Muy útil", "Algo útil"])
        await page.get_by_label("¿Qué tan útil te resulta la retroalimentación que el tutor virtual ofrece cuando aciertas o fallas una pregunta?")\
            .locator("..").get_by_text(opcion_p11, exact=True).click()
            

        # Pregunta 12
        print("Respondiendo Pregunta 12...")
        opciones_12 = [
            "Busqué la respuesta en mis apuntes o libros.",
            "Consulté con un profesor.",
            "Pedí ayuda a un compañero.",
            "Usé otra herramienta en línea."
        ]

        seleccionadas = random.sample(opciones_12, k=random.randint(1, 3))
        for opcion in seleccionadas:
            await page.get_by_role("checkbox", name=opcion).check()
        # Pregunta 13
        print("Respondiendo Pregunta 13...")
        opciones_13 = [
            "Ha mejorado mucho",
            "Ha mejorado bastante",
            "Ha mejorado moderadamente"
        ]
        opcion_13 = random.choice(opciones_13)
        await page.get_by_label("¿Sientes que con el uso del Tutor Virtual ha mejorado tu comprensión general de los temas de Química?")\
            .locator("..").get_by_text(opcion_13, exact=True).click()

        # Pregunta 14
        print("Respondiendo Pregunta 14...")
        opciones_14 = ["Totalmente", "Bastante", "Moderadamente"]
        opcion_14 = random.choice(opciones_14)
        await page.get_by_label("¿Crees que el tutor virtual ha complementado de manera efectiva las explicaciones dadas en clase?")\
            .locator("..").get_by_text(opcion_14, exact=True).click()

        # Pregunta 15
        print("Respondiendo Pregunta 15...")
        estrella_15 = random.choice([3, 4, 5])
        await page.get_by_role("radiogroup", name="15. En una escala del 1 al 5").locator(f'[aria-label="{estrella_15} Star"]').click()

        # Pregunta 16
        print("Respondiendo Pregunta 16...")
        estrella_16 = random.choice([3, 4, 5])
        await page.get_by_role("radiogroup", name="16. En una escala del 1 al 5").locator(f'[aria-label="{estrella_16} Star"]').click()
        
        # Pregunta 17

        print("Respondiendo Pregunta 17...")
        opciones_17 = ["Muy satisfecho", "Satisfecho"]
        opcion_17 = random.choice(opciones_17)
        await page.get_by_label("¿Qué tan satisfecho estas con el tutor virtual?")\
            .locator("..").get_by_text(opcion_17, exact=True).click()
            
        # Pregunta 18
        print("Respondiendo Pregunta 18...")
        opciones_18 = ["No, todo ha sido fácil de usar", "Sí, algunas dificultades menores"]
        opcion_18 = random.choice(opciones_18)
        await page.get_by_label("¿Has encontrado alguna dificultad para usar el tutor virtual en estas ocho semanas?")\
            .locator("..").get_by_text(opcion_18, exact=True).click()

        # Pregunta 19
        print("Respondiendo Pregunta 19...")
        # Cargar respuestas posibles desde el archivo Excel
        df_19 = pd.read_excel("G:/My Drive/Data Analysis/Proyectos/UCSUR/Encuestas Greta/pregunta_19.xlsx")
        respuestas_disponibles = df_19.iloc[:, 0].dropna().tolist()
        # Elegir aleatoriamente si se responde o no (50%)
        valor_random_19 = random.random()
        if valor_random_19 < 0.3:
            print("No se responderá la Pregunta 19.")
        else:
            respuesta_19 = random.choice(respuestas_disponibles)
        await page.get_by_label("¿Qué tipo de dificultades haz encontrado?").fill(respuesta_19)
            
        # Pregunta 20
        print("Respondiendo Pregunta 20...")
        await page.get_by_label("¿Recomendarías el uso del Tutor Virtual a otros estudiantes del curso?")\
            .locator("..").get_by_text("Sí", exact=True).click()

        # Pregunta 21
        print("Respondiendo Pregunta 21...")
        df_21 = pd.read_excel("G:/My Drive/Data Analysis/Proyectos/UCSUR/Encuestas Greta/pregunta_21.xlsx")
        respuestas_disponibles = df_21.iloc[:, 0].dropna().tolist()
        valor_random_21 = random.random()

        if valor_random_21 < 0.5:
            print("No se responderá la Pregunta 21.")
        else:
            respuesta_21 = random.choice(respuestas_disponibles)
            await page.wait_for_timeout(1000)  # Espera adicional por seguridad
            await page.get_by_placeholder("Enter your answer").nth(1).fill(respuesta_21)
        # Pregunta 22
        print("Respondiendo Pregunta 22...")

        # Cargar respuestas posibles desde el archivo Excel
        df_22 = pd.read_excel("G:/My Drive/Data Analysis/Proyectos/UCSUR/Encuestas Greta/pregunta_22.xlsx")
        respuestas_22_disponibles = df_22.iloc[:, 0].dropna().tolist()

        # Elegir aleatoriamente si se responde o no (50%)
        valor_random_22 = random.random()
        if valor_random_22 < 0.3:
            print("No se responderá la Pregunta 22.")
        else:
            respuesta_22 = random.choice(respuestas_22_disponibles)
            await page.get_by_label("¿Qué características adicionales te gustaría que tuviera el Tutor Virtual?").fill(respuesta_22)

        # Pregunta 23
        print("Respondiendo Pregunta 23...")

        cursos_23 = [
            "Bioquímica",
            "Biología",
            "Física",
            "Matemática general",
            "Estadística general",
            "Álgebra",
            "Matemática I",
            "Matemática II",
            "Matemática III"
        ]

        valor_random_23 = random.random()

        if valor_random_23 < 0.3:
            print("No se responderá la Pregunta 23.")
        else:
            respuesta_23 = "Sí, para " + random.choice(cursos_23)
            await page.get_by_label("¿Consideras que sería útil utilizar esta herramienta en otros cursos? Si es así, ¿en cuáles?").fill(respuesta_23)

        print("Enviando formulario...")
        await page.get_by_role("button", name="Submit").click()
        await page.wait_for_timeout(2000)  # Espera para asegurar el envío


        #Captura       
        await page.wait_for_timeout(1000)  # espera opcional para asegurar carga
        await page.screenshot(path=r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Encuestas Greta\test.png", full_page=True)
        await browser.close()

asyncio.run(run())
