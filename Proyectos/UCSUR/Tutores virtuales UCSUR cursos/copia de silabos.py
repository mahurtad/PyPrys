import os
import shutil

# Rutas
ruta_carreras = r"C:\Users\manue\OneDrive - Grupo Educad\Eventos de Innovación Educativa Docente - Materiales - tutores virtuales CCBB"
ruta_destino = r"C:\Users\manue\Downloads\Silabos\Silabos"

# Crear carpeta destino si no existe
os.makedirs(ruta_destino, exist_ok=True)

# Recorrer todas las carpetas dentro de la ruta base
for carpeta in os.listdir(ruta_carreras):
    ruta_carpeta = os.path.join(ruta_carreras, carpeta)

    if not os.path.isdir(ruta_carpeta):
        continue

    # Buscar subcarpeta "Sílabo"
    ruta_silabo = os.path.join(ruta_carpeta, "Sílabo")
    if os.path.exists(ruta_silabo):
        for archivo in os.listdir(ruta_silabo):
            ruta_archivo = os.path.join(ruta_silabo, archivo)
            if os.path.isfile(ruta_archivo):
                shutil.copy(ruta_archivo, ruta_destino)

print("Archivos copiados exitosamente.")
