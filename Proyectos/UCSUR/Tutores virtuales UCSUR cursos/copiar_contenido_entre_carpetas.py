import os
import shutil

# Rutas de origen y destino
ruta_origen = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"
ruta_destino_base = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\cursos"

# Lista de carpetas de interés
nombres_cursos = [
    "Biologia",
    "Bioquimica",
    "Estadistica General",
    "Fisica",
    "Fisica I",
    "Quimica general",
    "Química Orgánica"
]

# Recorremos la estructura de carpetas y archivos en la ruta de origen
for carpeta_actual, subcarpetas, archivos in os.walk(ruta_origen):
    for nombre_curso in nombres_cursos:
        if nombre_curso.lower() in carpeta_actual.lower():
            destino = os.path.join(ruta_destino_base, nombre_curso)
            os.makedirs(destino, exist_ok=True)

            for archivo in archivos:
                origen_archivo = os.path.join(carpeta_actual, archivo)
                destino_archivo = os.path.join(destino, archivo)

                # Evita sobrescribir si el archivo ya existe
                if not os.path.exists(destino_archivo):
                    shutil.copy2(origen_archivo, destino_archivo)
                    print(f"Copiado: {origen_archivo} -> {destino_archivo}")
                else:
                    print(f"Ya existe: {destino_archivo}")

print("Proceso finalizado.")
