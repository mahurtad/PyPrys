import os

# Ruta base
base_path = r'C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets\Finalizados\2025'

# Mapeo de nombres de carpetas a nombres base de archivos
target_folders = {
    '4. Conformidad Usuario': 'RE 01.1-CU-Conformidad de Usuario',
    '5. Conformidad Jefe Usuario': 'RE 01.2-CJ-Conformidad de Jefe de Usuario'
}

# Recorrer todos los directorios
for dirpath, dirnames, filenames in os.walk(base_path):
    folder_name = os.path.basename(dirpath)
    if folder_name in target_folders:
        new_base_name = target_folders[folder_name]
        for file in filenames:
            file_path = os.path.join(dirpath, file)
            file_ext = os.path.splitext(file)[1]  # extensión (por ejemplo .pdf o .docx)

            # Nuevo nombre completo con extensión
            new_file_name = f"{new_base_name}{file_ext}"
            new_file_path = os.path.join(dirpath, new_file_name)

            # Solo renombrar si el nombre actual es distinto
            if file_path != new_file_path:
                try:
                    os.rename(file_path, new_file_path)
                    print(f"Renombrado: {file} -> {new_file_name}")
                except Exception as e:
                    print(f"Error al renombrar {file}: {e}")
            else:
                print(f"Ya estaba correctamente nombrado: {file}")
