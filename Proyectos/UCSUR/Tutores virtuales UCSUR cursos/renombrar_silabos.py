import os

# Ruta de los archivos Word
ruta_words = r"C:\Users\manue\Downloads\Silabos\Word"

# Recorrer todos los archivos .docx en la carpeta
for archivo in os.listdir(ruta_words):
    if archivo.lower().endswith(".docx"):
        ruta_original = os.path.join(ruta_words, archivo)

        # Obtener el nombre del curso sin la extensión
        nombre_curso = os.path.splitext(archivo)[0].strip()

        # Crear nuevo nombre en formato "SILABO - CURSO"
        nuevo_nombre = f"SILABO - {nombre_curso}".upper() + ".docx"
        ruta_nueva = os.path.join(ruta_words, nuevo_nombre)

        # Renombrar archivo
        os.rename(ruta_original, ruta_nueva)

print("Renombramiento completado.")
