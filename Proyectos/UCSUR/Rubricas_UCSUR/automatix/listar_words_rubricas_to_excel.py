import os
import pandas as pd

# Rutas
ruta_base = r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\Reporte"
archivo_reporte = os.path.join(ruta_base, "documentos_word.xlsx")
archivo_manuel = os.path.join(ruta_base, "Lista de docentes y cursos - proyecto rúbricas 2025-2.xlsx")

# Leer hoja "Resumen por Curso"
df_resumen = pd.read_excel(archivo_reporte, sheet_name="Resumen por Curso")

# Leer hoja "Manuel" del archivo de docentes
df_manuel = pd.read_excel(archivo_manuel, sheet_name="Manuel")
df_manuel.columns = df_manuel.columns.str.strip().str.lower()

# Detectar la columna que contiene 'curso'
columna_curso = next((col for col in df_manuel.columns if "curso" in col), None)
if columna_curso is None:
    raise ValueError("No se encontró una columna que contenga 'curso' en el archivo Manuel.")

# Obtener lista de cursos normalizados
cursos_filtrar = df_manuel[columna_curso].dropna().astype(str).str.lower().tolist()

# Filtrar cursos del resumen por coincidencia contextual
filtrados = []
for _, row in df_resumen.iterrows():
    curso_resumen = str(row['Curso']).lower()
    if any(f in curso_resumen for f in cursos_filtrar):
        filtrados.append(row)

df_filtrado = pd.DataFrame(filtrados)

# Agregar como hoja nueva al Excel existente
with pd.ExcelWriter(archivo_reporte, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_filtrado.to_excel(writer, sheet_name='Manuel', index=False)

print(f"✅ Hoja 'Manuel' creada con {len(df_filtrado)} cursos coincidentes.")
