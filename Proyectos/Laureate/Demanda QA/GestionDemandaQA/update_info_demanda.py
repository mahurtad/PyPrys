import pandas as pd

# Rutas de los archivos
archivo_origen = r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones.xlsx"
hoja_origen = "Tickets Diarios"

archivo_actualizacion = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets\Tickets_Filtrados.xlsx"

# Leer los archivos
df_origen = pd.read_excel(archivo_origen, sheet_name=hoja_origen, dtype=str)
df_actualizacion = pd.read_excel(archivo_actualizacion, dtype=str)

# Asegurar que SCTASK esté como string y sin espacios
df_origen["SCTASK"] = df_origen["SCTASK"].str.strip()
df_actualizacion["SCTASK"] = df_actualizacion["SCTASK"].str.strip()

# Realizar la actualización usando merge
df_merged = df_origen.merge(
    df_actualizacion[["SCTASK", "CODIGO CHG", "ESTADO CHG"]],
    on="SCTASK",
    how="left",
    suffixes=("", "_actualizado")
)

# Actualizar columnas si hay datos nuevos
df_merged["CODIGO CHG"] = df_merged["CODIGO CHG_actualizado"].combine_first(df_merged["CODIGO CHG"])
df_merged["ESTADO CHG"] = df_merged["ESTADO CHG_actualizado"].combine_first(df_merged["ESTADO CHG"])

# Eliminar columnas auxiliares
df_merged.drop(columns=["CODIGO CHG_actualizado", "ESTADO CHG_actualizado"], inplace=True)

# Guardar los cambios sobrescribiendo el archivo original
with pd.ExcelWriter(archivo_origen, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    df_merged.to_excel(writer, sheet_name=hoja_origen, index=False)

print("Actualización completada.")
