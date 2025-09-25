import os
import shutil
import pandas as pd

# Ruta del archivo Excel con la hoja "Manuel"
archivo_excel = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\Reporte\documentos_word.xlsx"

# Leer la hoja "Manuel"
df = pd.read_excel(archivo_excel, sheet_name="Manuel")

# Ruta destino donde se copiarán las carpetas
destino_base = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\Reporte\Rúbricas extraídas"
os.makedirs(destino_base, exist_ok=True)

# Recorrer las rutas únicas desde "Ruta completa"
for ruta in df["Ruta completa"].dropna().unique():
    if os.path.isdir(ruta):
        nombre_carpeta = os.path.basename(ruta)
        destino = os.path.join(destino_base, nombre_carpeta)

        if not os.path.exists(destino):
            shutil.copytree(ruta, destino)
            print(f"✅ Copiado: {ruta} → {destino}")
        else:
            print(f"⚠️ Ya existe: {destino}")
    else:
        print(f"❌ No existe la carpeta: {ruta}")
