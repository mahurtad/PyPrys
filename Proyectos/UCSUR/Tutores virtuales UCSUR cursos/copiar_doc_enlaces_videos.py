import os
import shutil

# Ruta de origen (donde están los materiales originales)
ruta_origen = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"

# Ruta destino (raíz donde van los archivos)
ruta_destino_base = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\cursos"

# Cursos a procesar
nombres_cursos = [
    "Biologia",
    "Bioquimica",
    "Estadistica General",
    "Fisica",
    "Fisica I",
    "Quimica general",
    "Química Orgánica"
]

# Subcarpeta donde se encuentra el archivo en el origen
nombre_subcarpeta_videos = "Videos del curso"
nombre_base_archivo = "Enlaces de Videos"

# Recorremos cada curso
for curso in nombres_cursos:
    # Ruta origen esperada del archivo
    carpeta_curso_origen = None

    # Buscar la ruta exacta en el origen que coincida con el nombre del curso
    for carpeta_actual, subcarpetas, _ in os.walk(ruta_origen):
        if curso.lower() in carpeta_actual.lower():
            carpeta_curso_origen = os.path.join(carpeta_actual, nombre_subcarpeta_videos)
            break

    if carpeta_curso_origen and os.path.exists(carpeta_curso_origen):
        # Buscar el archivo dentro de la subcarpeta
        encontrado = False
        for archivo in os.listdir(carpeta_curso_origen):
            if archivo.startswith(nombre_base_archivo):
                origen_archivo = os.path.join(carpeta_curso_origen, archivo)
                destino_archivo = os.path.join(ruta_destino_base, curso, archivo)

                os.makedirs(os.path.join(ruta_destino_base, curso), exist_ok=True)
                shutil.copy2(origen_archivo, destino_archivo)
                print(f"Copiado: {origen_archivo} -> {destino_archivo}")
                encontrado = True
                break

        if not encontrado:
            print(f"No se encontró '{nombre_base_archivo}' en {carpeta_curso_origen}")
    else:
        print(f"No se encontró la subcarpeta 'Videos del curso' para el curso: {curso}")

print("Proceso finalizado.")
