import os
import re
from PyPDF2 import PdfReader

# Ruta donde se encuentran los PDFs
ruta_certificados = r'C:\Users\manue\Downloads\Generador\CAPACITACIÓN DISEÑO DE EXPERIENCIAS DE APRENDIZAJE CON INTELIGENCIA ARTIFICIAL\Certificados'

# Función para extraer el nombre del PDF
def extraer_nombre_pdf(ruta_pdf):
    lector_pdf = PdfReader(ruta_pdf)
    texto_completo = ""
    
    # Extraer texto de cada página del PDF
    for pagina in lector_pdf.pages:
        texto_completo += pagina.extract_text()
    
    # Aquí debes ajustar el patrón de búsqueda según cómo esté resaltado el nombre
    # Ejemplo de búsqueda de nombre completo con un patrón genérico: dos palabras que empiecen con mayúscula
    patron_nombre = re.compile(r'([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)')
    coincidencias = patron_nombre.findall(texto_completo)
    
    if coincidencias:
        # Retorna la primera coincidencia (asumiendo que es el nombre y apellido)
        return " ".join(coincidencias[0])
    
    return None

# Iterar sobre todos los archivos en la ruta y renombrar los PDFs
for archivo in os.listdir(ruta_certificados):
    if archivo.endswith('.pdf'):
        ruta_pdf = os.path.join(ruta_certificados, archivo)
        nombre = extraer_nombre_pdf(ruta_pdf)
        
        if nombre:
            # Crear una nueva ruta con el nombre extraído
            nuevo_nombre = f"{nombre}.pdf"
            nueva_ruta_pdf = os.path.join(ruta_certificados, nuevo_nombre)
            
            # Renombrar el archivo
            os.rename(ruta_pdf, nueva_ruta_pdf)
            print(f"Renombrado: {archivo} -> {nuevo_nombre}")
        else:
            print(f"No se encontró el nombre en {archivo}")
