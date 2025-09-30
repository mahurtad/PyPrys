# analisis_tickets_visual.py
import pandas as pd
import numpy as np

# ---------------- CONFIG ----------------
EXCEL_PATH   = r"C:\Users\manue\OneDrive - EduCorpPERU\Calidad de Software - Certificaciones\Gestión Demanda Certificaciones LIUv1.xlsx"
EXCEL_SHEET  = "Tickets Diarios"

# Si querés guardar resultados en archivos, pon True; por defecto solo visual.
SAVE = False
OUT_DIR = "."


# --------------- LECTURA ---------------
df = pd.read_excel(EXCEL_PATH, sheet_name=EXCEL_SHEET, engine="openpyxl")
df.columns = df.columns.str.strip()

# --------------- LIMPIEZA INICIAL ---------------
first_col = df.columns[0]
mask_repeated_header = df[first_col].astype(str).str.strip().eq(first_col.strip())
mask_all_na = df.isna().all(axis=1)
df_clean = df[~mask_repeated_header & ~mask_all_na].copy()
df_clean.reset_index(drop=True, inplace=True)

# --------------- DETECCIÓN DE COLUMNAS (flexible) ---------------
col_institucion = None
col_ritm = None
col_analista = None

for c in df_clean.columns:
    cu = c.upper().replace(" ", "")
    if "INSTITUCION" in cu or "INSTITUCIÓN" in c.upper() or "INSTITUT" in cu:
        col_institucion = col_institucion or c
    if "RITM" in cu or "INC" in cu:
        col_ritm = col_ritm or c
    if "ANALISTA" in c.upper() and "QA" in c.upper():
        col_analista = col_analista or c

# heurística adicional
if col_institucion is None:
    for c in df_clean.columns:
        if "INSTIT" in c.upper():
            col_institucion = c
            break
if col_ritm is None:
    for c in df_clean.columns:
        if "RITM" in c.upper() or "INC" in c.upper():
            col_ritm = c
            break
if col_analista is None:
    for c in df_clean.columns:
        if "ANALISTA" in c.upper():
            col_analista = c
            break

missing = [name for name, col in [("INSTITUCION", col_institucion), ("RITM/INC", col_ritm), ("ANALISTA QA", col_analista)] if col is None]
if missing:
    raise SystemExit(f"ERROR: faltan columnas: {', '.join(missing)}. Columnas detectadas: {list(df_clean.columns)}")

# --------------- FILTRAR REGISTROS SIN INSTITUCION ---------------
df_clean[col_institucion] = df_clean[col_institucion].astype(str).str.strip()
df_clean.loc[df_clean[col_institucion].str.lower() == "nan", col_institucion] = np.nan
mask_missing_institucion = df_clean[col_institucion].isna() | (df_clean[col_institucion] == "")
df_filtered = df_clean[~mask_missing_institucion].copy()
df_filtered.reset_index(drop=True, inplace=True)

# Normalizaciones mínimas
df_filtered[col_institucion] = df_filtered[col_institucion].astype(str).str.strip()
df_filtered[col_analista] = df_filtered[col_analista].astype(str).str.strip()
df_filtered[col_ritm] = df_filtered[col_ritm].apply(lambda x: x.strip() if isinstance(x, str) else x)

# --------------- CÁLCULOS ---------------
total_filas = len(df_filtered)

# Por institución
conteo_institucion = df_filtered.groupby(col_institucion).size().reset_index(name="total_filas").sort_values("total_filas", ascending=False)
filas_with_ritm_by_inst = df_filtered[df_filtered[col_ritm].notna()].groupby(col_institucion).size().reset_index(name="filas_con_ritm")
unique_ritm_by_inst = df_filtered.dropna(subset=[col_ritm]).groupby(col_institucion)[col_ritm].nunique().reset_index(name="ritm_unicos")

summary_inst = (conteo_institucion
                .merge(filas_with_ritm_by_inst, on=col_institucion, how="left")
                .merge(unique_ritm_by_inst, on=col_institucion, how="left")
                .fillna(0))
summary_inst["filas_con_ritm"] = summary_inst["filas_con_ritm"].astype(int)
summary_inst["ritm_unicos"] = summary_inst["ritm_unicos"].astype(int)
summary_inst["percent_total(%)"] = (summary_inst["total_filas"] / total_filas * 100).round(2)

# Por analista (sobre los datos filtrados)
conteo_analista = df_filtered.groupby(col_analista).size().reset_index(name="conteo_filas").sort_values("conteo_filas", ascending=False)
unique_ritm_analista = df_filtered.dropna(subset=[col_ritm]).groupby(col_analista)[col_ritm].nunique().reset_index(name="ritm_unicos")
summary_analista = conteo_analista.merge(unique_ritm_analista, on=col_analista, how="outer").fillna(0)
summary_analista["conteo_filas"] = summary_analista["conteo_filas"].astype(int)
summary_analista["ritm_unicos"] = summary_analista["ritm_unicos"].astype(int)

# Pivot Analista x Institución (conteo de filas)
pivot_counts = pd.crosstab(df_filtered[col_analista], df_filtered[col_institucion])

# Pivot Analista x Institución (RITM únicos)
pivot_ritm_unicos = df_filtered.dropna(subset=[col_ritm]).pivot_table(index=col_analista,
                                                                        columns=col_institucion,
                                                                        values=col_ritm,
                                                                        aggfunc=pd.Series.nunique,
                                                                        fill_value=0)

# --------------- SALIDA VISUAL (solo en consola) ---------------
# Resumen general (números clave)
print("\nResumen (solo registros con INSTITUCION):")
print(f"Total filas consideradas: {total_filas}")
print(f"Registros excluidos por no tener INSTITUCION: {mask_missing_institucion.sum()}")
print("-" * 60)

# Top instituciones (visual)
print("\nTop instituciones (por total de filas):")
print(summary_inst.head(20).to_string(index=False))

# Top analistas (visual)
print("\nTop analistas (por conteo de filas):")
print(summary_analista.head(20).to_string(index=False))

# Mostrar porción del pivot (para no saturar la consola)
print("\nPivot (Analista x Institución) — muestra filas top 10 y primeras 10 columnas:")
if not pivot_counts.empty:
    # seleccionar top 10 analistas por suma de sus filas
    top_analistas = pivot_counts.sum(axis=1).sort_values(ascending=False).head(10).index
    cols_first10 = pivot_counts.columns[:10]
    print(pivot_counts.loc[top_analistas, cols_first10].to_string())
else:
    print("No hay datos para pivotar.")

print("\nListo ✅ (salida visual).")
# --------------- (Opcional) guardar si SAVE=True ---------------
if SAVE:
    import os
    os.makedirs(OUT_DIR, exist_ok=True)
    summary_inst.to_csv(f"{OUT_DIR}/resumen_conteos_por_institucion_filtrado.csv", index=False)
    summary_analista.to_csv(f"{OUT_DIR}/resumen_conteos_por_analista_filtrado.csv", index=False)
    pivot_counts.to_csv(f"{OUT_DIR}/pivot_analista_institucion_counts_filtrado.csv")
    pivot_ritm_unicos.to_csv(f"{OUT_DIR}/pivot_analista_institucion_ritm_unicos_filtrado.csv")
