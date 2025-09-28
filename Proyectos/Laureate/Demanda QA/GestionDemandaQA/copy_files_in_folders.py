import os
import shutil
import zipfile

# Definir las rutas de origen y destino
source_path = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets\XXX-12092025-RITM0555039-SCTASK1412640-IntDocAdm-Eliminar solicitudes duplicadas en PRACTICA_PREPRO-MAHURTAD"
destination_path = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets\Upload_to_SN"

# Extraer el nombre del tercer segmento de la carpeta de origen
folder_name = os.path.basename(source_path)  # Obtiene el último segmento de la ruta
segments = folder_name.split('-')  # Divide el nombre en segmentos usando '-'

if len(segments) >= 3:
    folder_name_segment = segments[2]  # Tomar el tercer segmento como nombre
else:
    folder_name_segment = "backup"  # Nombre por defecto si no se encuentra el segmento esperado

# Crear la nueva carpeta dentro de la ruta destino
new_folder_path = os.path.join(destination_path, folder_name_segment)
os.makedirs(new_folder_path, exist_ok=True)

# Copiar los archivos manejando duplicados
file_names_seen = set()  # Para rastrear nombres de archivo únicos
for root, _, files in os.walk(source_path):
    for file in files:
        source_file = os.path.join(root, file)
        
        # Comprobar si el archivo ya existe en la carpeta destino
        destination_file = os.path.join(new_folder_path, file)
        file_base, file_extension = os.path.splitext(file)  # Separar nombre y extensión
        counter = 1

        while file in file_names_seen or os.path.exists(destination_file):
            # Renombrar el archivo si ya existe
            file = f"{file_base}_{counter}{file_extension}"
            destination_file = os.path.join(new_folder_path, file)
            counter += 1

        # Copiar el archivo al destino
        shutil.copy2(source_file, destination_file)
        file_names_seen.add(file)  # Registrar el archivo copiado
        print(f"Copiado: {source_file} -> {destination_file}")

print(f"Archivos copiados a la carpeta: {new_folder_path}")


