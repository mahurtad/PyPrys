from PyPDF2 import PdfMerger
from fpdf import FPDF
import os

# Ruta donde están los archivos PDF por carrera
ruta_base = r"G:\My Drive\Shared GyM\UCSUR\Internacionalización\guías para cada carrera 20252\Guias de intercambio 2025-2"

# Obtener lista de archivos PDF
archivos = sorted([
    f for f in os.listdir(ruta_base)
    if f.lower().endswith(".pdf")
])

# Función para limpiar texto y crear portada en PDF
def crear_portada(texto, nombre_archivo):
    texto_limpio = (
        texto.replace("–", "-")  # guion largo a simple
             .encode('latin-1', errors='ignore')  # eliminar caracteres no compatibles
             .decode('latin-1')
    )
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", style='B', size=16)
    pdf.multi_cell(0, 10, f"Guía de Intercambio Internacional 2025-2\n{texto_limpio}", align='C')
    path = os.path.join(ruta_base, nombre_archivo)
    pdf.output(path)
    return path

# Inicializar el PDF final
merger = PdfMerger()

# Recorrer cada guía
for archivo in archivos:
    carrera = archivo.replace("Guía Intercambio Internacional 2025-2 – ", "").replace(".pdf", "")
    nombre_portada = f"portada_{carrera}.pdf"
    
    # Crear portada y agregar ambas partes al merger
    path_portada = crear_portada(carrera, nombre_portada)
    merger.append(path_portada)
    merger.append(os.path.join(ruta_base, archivo))

# Guardar el resultado
pdf_salida = os.path.join(ruta_base, "Guías_Carreras_Internacional_2025-2.pdf")
merger.write(pdf_salida)
merger.close()

print(f"✅ PDF consolidado generado correctamente en:\n{pdf_salida}")
