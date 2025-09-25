import os
import pytesseract
from PIL import Image
import pandas as pd

# Ruta de Tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Carpeta donde están todas las imágenes capturadas
carpeta_imagenes = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\Silabo\Capturas\QUÍMICA ORGÁNICA"

# Preparar contenedor de datos
datos = []

# Procesar todas las imágenes en orden
for archivo in sorted(os.listdir(carpeta_imagenes)):
    if archivo.endswith(".png"):
        ruta = os.path.join(carpeta_imagenes, archivo)
        imagen = Image.open(ruta)
        texto = pytesseract.image_to_string(imagen, lang="spa")

        # Dividir el texto entre columnas clave
        partes = texto.split("Tipo de actividad:")
        if len(partes) >= 3:
            docente_raw = partes[1].strip()
            autonomo_raw = partes[2].strip()
        else:
            docente_raw = partes[1].strip() if len(partes) > 1 else ""
            autonomo_raw = ""

        datos.append({
            "ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE": docente_raw,
            "ACTIVIDADES DE TRABAJO AUTÓNOMO": autonomo_raw
        })

# Exportar a Excel
df = pd.DataFrame(datos)
salida = os.path.join(carpeta_imagenes, "Actividades_Extraidas_Desde_OCR.xlsx")
df.to_excel(salida, index=False)

print(f"✅ Archivo generado: {salida}")
