# excel_probe.py
import pandas as pd
from pathlib import Path

EXCEL_PATH  = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET       = 0  # o el nombre exacto de la hoja si lo conoces

CAND_PIDM = ["PIDM", "pidm", "Pidm", "PIDM "]
CAND_PROG = ["COD_PROGRAMA", "codigoPrograma", "CodigoPrograma", "cod_programa", "COD_PROGRAMA "]

def find_col(cols, candidates):
    # busca la primera coincidencia ignorando espacios y mayúsculas
    norm = {c.strip().lower(): c for c in cols}
    for cand in candidates:
        k = cand.strip().lower()
        if k in norm:
            return norm[k]
    return None

p = Path(EXCEL_PATH)
print(f"[INFO] Leyendo: {p}  (exists={p.exists()})")

df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET, dtype=str)
print(f"[INFO] Columnas: {list(df.columns)}")

pidm_col = find_col(df.columns, CAND_PIDM)
prog_col = find_col(df.columns, CAND_PROG)
print(f"[INFO] Detectadas → PIDM={pidm_col}  COD_PROGRAMA={prog_col}")

if not pidm_col or not prog_col:
    raise SystemExit("[ERROR] No encontré columnas PIDM / COD_PROGRAMA (o equivalentes). Revisa encabezados.")

df = df[[pidm_col, prog_col]].dropna()
df[pidm_col] = df[pidm_col].astype(str).str.extract(r"(\d+)", expand=False)
df[prog_col] = df[prog_col].astype(str).str.strip()
df = df[(df[pidm_col].str.len() > 0) & (df[prog_col].str.len() > 0)].drop_duplicates([pidm_col, prog_col])

pairs = list(df.apply(lambda r: (r[prog_col], r[pidm_col]), axis=1))
print(f"[RESULT] Pairs cargados: {len(pairs)}")
print("[RESULT] Primeros 5:", pairs[:5])
