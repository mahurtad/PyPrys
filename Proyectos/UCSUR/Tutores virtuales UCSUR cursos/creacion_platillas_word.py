from docx import Document
import os

# Lista de cursos
cursos = [
    "Biología",
    "Bioquímica",
    "Estadística General",
    "Física",
    "Física I",
    "Química General",
    "Química Orgánica"
]

# Ruta donde se guardarán los archivos
ruta_destino = r"C:\Users\manue\Downloads\Silabos"

# Crear archivos Word
for curso in cursos:
    nombre_archivo = f"SILABO - {curso}".upper() + ".docx"
    ruta_completa = os.path.join(ruta_destino, nombre_archivo)
    
    doc = Document()
    doc.add_heading(f"Sílabo del curso: {curso}", level=1)
    doc.save(ruta_completa)

print("Archivos creados correctamente.")
