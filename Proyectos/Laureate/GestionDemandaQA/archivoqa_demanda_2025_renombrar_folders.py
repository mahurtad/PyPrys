import os
import pandas as pd
import re

# Ruta de carpetas a analizar
folder_path = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets"
# Ruta del archivo adicional con datos de "NUM."
certificaciones_path = r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones.xlsx"

# Leer la hoja "Tickets Diarios"
certificaciones_df = pd.read_excel(certificaciones_path, sheet_name="Tickets Diarios")

# Recorrer las carpetas en la ruta
for folder_name in os.listdir(folder_path):
    full_path = os.path.join(folder_path, folder_name)

    # Verificar si es una carpeta
    if os.path.isdir(full_path):
        # Dividir el nombre de la carpeta en fragmentos separados por "-"
        fragments = folder_name.split("-")

        # Asegurar que haya al menos tres fragmentos para extraer el tercer segmento (RITM/INC)
        if len(fragments) > 2:
            ritm_inc = fragments[2]  # Extraer el código "RITMxxxxxx" o "INCxxxxxx"
        else:
            continue  # Si no hay suficientes fragmentos, pasar a la siguiente carpeta

        # Buscar el número correspondiente en "NUM." según "RITM/INC"
        num = certificaciones_df.loc[
            certificaciones_df["RITM/INC"] == ritm_inc, "NUM."
        ].values
        num = str(int(num[0])) if len(num) > 0 else "XXX"  # Si no encuentra, pone "SIN_NUM"

        # Si hay más fragmentos, reemplazar el primer segmento con el número encontrado
        if len(fragments) > 0:
            old_name = folder_name  # Guardamos el nombre original
            fragments[0] = num  # Reemplaza "XXX" con el número encontrado
            new_folder_name = "-".join(fragments)

            # Rutas completa de la carpeta antes y después
            new_full_path = os.path.join(folder_path, new_folder_name)

            # Renombrar la carpeta solo si el nuevo nombre no existe ya
            if not os.path.exists(new_full_path):
                os.rename(full_path, new_full_path)
                print(f"✔ Carpeta renombrada: {old_name} → {new_folder_name}")
            else:
                print(f"⚠ Carpeta no renombrada (ya existe): {new_folder_name}")
