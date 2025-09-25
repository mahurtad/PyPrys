import os
import shutil
import zipfile

# Rutas base
ruta_origen = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"
ruta_destino_base = r"G:\My Drive\Shared GyM\UCSUR\Archivos para tutores virtuales\cursos"

# Lista de cursos
nombres_cursos = [
    "Biologia",
    "Bioquimica",
    "Estadistica General",
    "Fisica",
    "Fisica I",
    "Quimica general",
    "Química Orgánica"
]

# Subcarpetas a copiar y comprimir
subcarpetas_objetivo = [
    "Presentaciones",
    "Materiales adicionales",
    "Lecturas o bibliografía obligatoria"
]

# Función para comprimir una carpeta
def comprimir_carpeta(carpeta_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, _, files in os.walk(carpeta_path):
            for file in files:
                archivo_path = os.path.join(root, file)
                zipf.write(archivo_path, os.path.relpath(archivo_path, start=carpeta_path))

# Recorremos la estructura del origen
for carpeta_actual, subcarpetas, _ in os.walk(ruta_origen):
    for nombre_curso in nombres_cursos:
        if nombre_curso.lower() in carpeta_actual.lower():
            ruta_destino_curso = os.path.join(ruta_destino_base, nombre_curso)
            os.makedirs(ruta_destino_curso, exist_ok=True)

            for subcarpeta in subcarpetas:
                if subcarpeta in subcarpetas_objetivo:
                    origen_subcarpeta = os.path.join(carpeta_actual, subcarpeta)
                    destino_subcarpeta = os.path.join(ruta_destino_curso, subcarpeta)

                    # Copiar subcarpeta
                    if not os.path.exists(destino_subcarpeta):
                        shutil.copytree(origen_subcarpeta, destino_subcarpeta)
                        print(f"Copiada: {origen_subcarpeta} -> {destino_subcarpeta}")

                        # Crear archivo ZIP
                        zip_destino = os.path.join(ruta_destino_curso, f"{subcarpeta}.zip")
                        comprimir_carpeta(destino_subcarpeta, zip_destino)
                        print(f"Zip creado: {zip_destino}")

                        # Eliminar carpeta original
                        shutil.rmtree(destino_subcarpeta)
                        print(f"Eliminada carpeta: {destino_subcarpeta}")
                    else:
                        print(f"Ya existe: {destino_subcarpeta}")

print("Proceso finalizado.")
