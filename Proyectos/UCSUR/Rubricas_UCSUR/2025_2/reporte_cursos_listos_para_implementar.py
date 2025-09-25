# -*- coding: utf-8 -*-
"""
Genera un reporte por curso indicando si está listo para implementar (todos sus registros en "Validado"),
el conteo de actividades y la lista de evaluaciones (p. ej., EC1, EC2, EC3, EC4, ED, EF, EP).

Requisitos:
    pip install pandas openpyxl
"""
from pathlib import Path
import pandas as pd
import numpy as np
import sys

# Ruta fija del archivo Excel de entrada
INPUT_FILE = Path(r"G:\My Drive\Data Analysis\Proyectos\UCSUR\Rubricas_UCSUR\Reporte\Lista de docentes y cursos - proyecto rúbricas 2025-2_reporte.xlsx")
SHEET_NAME = "Base de datos"  # nombre de la hoja
# Ruta de salida (misma carpeta que el archivo de entrada)
OUTPUT_FILE = INPUT_FILE.parent / "reporte_cursos_validado.xlsx"

# Posibles nombres de columnas
COLUMNS_CANDIDATES = {
    "carrera": [
        "carrera del plan", "carrera", "carrera_plan", "carrera deL plan"
    ],
    "curso": [
        "nombre del curso", "curso", "asignatura", "nombre_curso"
    ],
    "estatus": [
        "estatus consultora", "estatus", "estado consultora", "estatus_consultora"
    ],
    "evaluacion": [
        "evaluación", "evaluaciones", "actividad", "tipo de evaluación",
        "tipo de evaluacion", "actividad/evaluación", "actividad/evaluacion",
        "evaluacion"
    ]
}

EVAL_ORDER = ["EC1", "EC2", "EC3", "EC4", "ED", "EF", "EP"]

def normalize_colnames(cols):
    def norm(s):
        return (
            str(s)
            .strip()
            .replace("\n", " ")
            .replace("\r", " ")
            .replace("  ", " ")
            .lower()
        )
    return [norm(c) for c in cols]

def find_column(df, candidates):
    cols_norm = normalize_colnames(df.columns)
    for wanted in candidates:
        w_norm = wanted.strip().lower()
        for i, c in enumerate(cols_norm):
            if c == w_norm:
                return df.columns[i]
    for wanted in candidates:
        w_norm = wanted.strip().lower()
        for i, c in enumerate(cols_norm):
            if w_norm in c:
                return df.columns[i]
    return None

def sort_evals(evals):
    primary = [e for e in EVAL_ORDER if e in evals]
    others = sorted([e for e in evals if e not in EVAL_ORDER])
    return primary + others

def main():
    if not INPUT_FILE.exists():
        print(f"[ERROR] No se encontró el archivo: {INPUT_FILE}")
        sys.exit(1)

    try:
        df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"[ERROR] No se pudo leer la hoja '{SHEET_NAME}': {e}")
        sys.exit(1)

    col_carrera = find_column(df, COLUMNS_CANDIDATES["carrera"])
    col_curso   = find_column(df, COLUMNS_CANDIDATES["curso"])
    col_estatus = find_column(df, COLUMNS_CANDIDATES["estatus"])
    col_eval    = find_column(df, COLUMNS_CANDIDATES["evaluacion"])

    missing = [name for name, col in [("Carrera del plan", col_carrera),
                                      ("Nombre del curso", col_curso),
                                      ("Estatus consultora", col_estatus)] if col is None]
    if missing:
        print(f"[ERROR] No se encontraron las columnas requeridas: {', '.join(missing)}")
        sys.exit(1)

    if col_eval is None:
        df["_eval_"] = np.nan
        col_eval = "_eval_"

    def to_text(x):
        if pd.isna(x):
            return ""
        return str(x).strip()

    df["_carrera_"] = df[col_carrera].apply(to_text)
    df["_curso_"]   = df[col_curso].apply(to_text)
    df["_estatus_"] = df[col_estatus].astype(str).str.strip().str.lower()
    df["_validado_"] = df["_estatus_"].eq("validado")
    df["_evalcode_"] = df[col_eval].apply(to_text).str.upper().str.replace(" ", "", regex=False)

    grouped = df.groupby(["_carrera_", "_curso_"], dropna=False)

    rows = []
    for (carrera, curso), g in grouped:
        all_valid = bool(g["_validado_"].all())
        evals = [e for e in g["_evalcode_"].unique().tolist() if e]
        evals_sorted = sort_evals(evals)
        evaluaciones = ", ".join(evals_sorted)
        rows.append({
            "Carrera del plan": carrera,
            "Nombre del curso": curso,
            "Conteo de actividades": len(evals_sorted),
            "Evaluaciones": evaluaciones,
            "Estado": "OK" if all_valid else "NO"
        })

    rep = pd.DataFrame(rows).sort_values(["Carrera del plan", "Nombre del curso"]).reset_index(drop=True)

    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as wr:
            rep.to_excel(wr, sheet_name="Reporte", index=False)
    except Exception as e:
        print(f"[ERROR] No se pudo escribir el archivo de salida: {e}")
        sys.exit(1)

    total = len(rep)
    ok = int((rep["Estado"] == "OK").sum())
    print(f"[OK] Reporte generado: {OUTPUT_FILE}")
    print(f"Total cursos: {total} | Listos para implementar (OK): {ok} | No listos: {total - ok}")

if __name__ == "__main__":
    main()
