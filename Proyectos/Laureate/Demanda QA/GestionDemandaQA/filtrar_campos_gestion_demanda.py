import pandas as pd
import os

# Ruta del archivo fuente
archivo_origen = r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones.xlsx"
hoja = "Tickets Diarios"

# Ruta de exportación
ruta_exportacion = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets"
nombre_archivo_salida = "Tickets_Filtrados.xlsx"
ruta_completa_salida = os.path.join(ruta_exportacion, nombre_archivo_salida)

# Leer el archivo Excel
df = pd.read_excel(archivo_origen, sheet_name=hoja)

# Filtrar por "ANALISTAS QA" = "Manuel Hurtado"
df_filtrado = df[df["ANALISTAS QA"] == "Manuel Hurtado"]

# Filtrar por "CODIGO CHG" vacío o nulo
df_filtrado = df_filtrado[df_filtrado["CODIGO CHG"].isna()]

# Mostrar valores de "RITM/INC" y "SCTASK" en consola
print("Valores filtrados de RITM/INC y SCTASK:")
print(df_filtrado[["RITM/INC", "SCTASK"]])

# Mostrar total de registros
print(f"\nTotal de registros encontrados: {len(df_filtrado)}")

# Seleccionar solo las columnas requeridas para exportación
columnas_a_exportar = [
    "NUM.",
    "RITM/INC",
    "SCTASK",
    "ANALISTAS QA",
    "CODIGO CHG",
    "ESTADO CHG"
]
df_exportar = df_filtrado[columnas_a_exportar]

# Exportar a archivo Excel
df_exportar.to_excel(ruta_completa_salida, index=False)
print(f"\nArchivo exportado exitosamente a: {ruta_completa_salida}")
