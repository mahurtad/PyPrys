import pandas as pd
import os

# Ruta base
ruta_base = r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets"

# Archivos
archivo_filtrado = os.path.join(ruta_base, "Tickets_Filtrados.xlsx")
archivo_exportado = os.path.join(ruta_base, "Exported_Tickets.xlsx")

# Cargar archivos
df_filtrado = pd.read_excel(archivo_filtrado)
df_exportado = pd.read_excel(archivo_exportado)

# Realizar merge usando SCTASK como clave
df_actualizado = pd.merge(
    df_filtrado,
    df_exportado[["SCTASK", "CHG", "State"]],
    on="SCTASK",
    how="left"
)

# Actualizar columnas CODIGO CHG y ESTADO CHG
df_actualizado["CODIGO CHG"] = df_actualizado["CHG"]
df_actualizado["ESTADO CHG"] = df_actualizado["State"]

# Eliminar columnas auxiliares si se desea
df_actualizado.drop(columns=["CHG", "State"], inplace=True)

# Guardar el archivo actualizado (sobrescribiendo el original)
df_actualizado.to_excel(archivo_filtrado, index=False)
print(f"Archivo actualizado exitosamente en: {archivo_filtrado}")
