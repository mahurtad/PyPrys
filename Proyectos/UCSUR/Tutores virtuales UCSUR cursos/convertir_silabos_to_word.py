import os
from pdf2docx import Converter

# Ruta donde están los PDFs
ruta_pdf = r"C:\Users\manue\Downloads\Silabos\Silabos"

# Recorrer todos los archivos en la ruta
for archivo in os.listdir(ruta_pdf):
    if archivo.lower().endswith(".pdf"):
        ruta_completa_pdf = os.path.join(ruta_pdf, archivo)
        nombre_sin_ext = os.path.splitext(archivo)[0]
        ruta_salida_docx = os.path.join(ruta_pdf, f"{nombre_sin_ext}.docx")
        
        # Convertir PDF a Word
        cv = Converter(ruta_completa_pdf)
        cv.convert(ruta_salida_docx)
        cv.close()

print("Conversión completada.")
