import os
import pandas as pd

# Ruta de carpetas a analizar
folder_path = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets"

# Ruta del archivo con datos
certificaciones_path = r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones.xlsx"

# Leer la hoja "Tickets Diarios"
certificaciones_df = pd.read_excel(certificaciones_path, sheet_name="Tickets Diarios")

# Recorrer las carpetas en la ruta
for folder_name in os.listdir(folder_path):
    full_path = os.path.join(folder_path, folder_name)

    # Verificar si es una carpeta
    if os.path.isdir(full_path):
        fragments = folder_name.split("-")

        # Asegurar que haya al menos cuatro fragmentos para extraer el SCTASK
        if len(fragments) > 3:
            sctask = fragments[3].strip()  # Extraer el código "SCTASKxxxxxxx"
        else:
            continue  # Saltar si no hay suficientes fragmentos

        # Buscar el número correspondiente en la columna "NUM." según el campo "SCTASK"
        num = certificaciones_df.loc[
            certificaciones_df["SCTASK"] == sctask, "NUM."
        ].values
        num = str(int(num[0])) if len(num) > 0 else "XXX"

        # Reemplazar el primer fragmento por el número encontrado
        old_name = folder_name
        fragments[0] = num
        new_folder_name = "-".join(fragments)

        new_full_path = os.path.join(folder_path, new_folder_name)

        if not os.path.exists(new_full_path):
            os.rename(full_path, new_full_path)
            print(f"✔ Carpeta renombrada: {old_name} → {new_folder_name}")
        else:
            print(f"⚠ Carpeta no renombrada (ya existe): {new_folder_name}")
