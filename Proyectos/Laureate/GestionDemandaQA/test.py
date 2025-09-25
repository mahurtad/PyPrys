import os
import sys
from pathlib import Path
import pandas as pd

# ----------------------------
# CONFIGURACIÓN
# ----------------------------
# 1) Ruta del Excel ADJUNTO (esta es la que pediste usar ahora)
EXCEL_PATH_ATTACHED = Path("/mnt/data/Gestión Demanda Certificaciones LIUv1.xlsx")

# 2) (Opcional) Ruta alternativa en tu PC Windows, por si luego lo ejecutas local
EXCEL_PATH_WINDOWS = Path(r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones LIUv1.xlsx")

# 3) Hoja a leer
SHEET_NAME = "Tickets Diarios"

# 4) Carpeta base donde están las carpetas a renombrar (ajústala si necesitas)
FOLDER_BASE = Path(r"C:\Users\manue\OneDrive - EduCorpPERU\2025\Tickets")

# 5) Nombres de columnas (tal como aparecen en tu hoja)
COL_SCTASK = "SCTASK"
COL_NUM = "NUM."   # si en tu archivo aparece sin punto, cámbialo a "NUM"

# ----------------------------
# UTILIDADES
# ----------------------------
def pick_excel_path() -> Path:
    """
    Prioriza el Excel adjunto en /mnt/data.
    Si no existe (por ejemplo, cuando lo corras en tu PC), usa la ruta de Windows.
    """
    if EXCEL_PATH_ATTACHED.exists():
        return EXCEL_PATH_ATTACHED
    if EXCEL_PATH_WINDOWS.exists():
        return EXCEL_PATH_WINDOWS
    raise FileNotFoundError(
        f"No se encontró el Excel ni en '{EXCEL_PATH_ATTACHED}' ni en '{EXCEL_PATH_WINDOWS}'."
    )

def load_mapping(excel_path: Path, sheet_name: str, col_sctask: str, col_num: str) -> pd.DataFrame:
    """
    Lee la hoja indicada y devuelve un DataFrame con SCTASK y NUM como texto, limpiando espacios.
    """
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, dtype={col_sctask: str, col_num: str})
    except ImportError as e:
        print("⚠ Falta el motor de Excel (openpyxl). Instálalo con: pip install openpyxl", file=sys.stderr)
        raise
    except ValueError as e:
        print(f"⚠ Verifica que exista la hoja '{sheet_name}' en '{excel_path.name}'.", file=sys.stderr)
        raise

    # Normalización básica
    if col_sctask not in df.columns or col_num not in df.columns:
        raise KeyError(
            f"Columnas requeridas no encontradas. Esperado: '{col_sctask}' y '{col_num}'. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    df[col_sctask] = df[col_sctask].astype(str).str.strip().str.upper()
    df[col_num] = df[col_num].astype(str).str.strip()

    # Elimina filas vacías en SCTASK
    df = df[df[col_sctask].notna() & df[col_sctask].ne("")].copy()

    # En caso de duplicados de SCTASK, nos quedamos con el primero
    df = df.drop_duplicates(subset=[col_sctask], keep="first").reset_index(drop=True)
    return df[[col_sctask, col_num]]

def find_num(mapping_df: pd.DataFrame, col_sctask: str, col_num: str, sctask_value: str) -> str:
    """
    Busca el NUM correspondiente al SCTASK (case-insensitive). Devuelve 'XXX' si no se encuentra.
    """
    key = (sctask_value or "").strip().upper()
    row = mapping_df.loc[mapping_df[col_sctask] == key]
    if not row.empty:
        value = str(row.iloc[0][col_num]).strip()
        # Limpieza típica cuando Excel deja 'nan' como texto
        if value.lower() in ("nan", "none"):
            return "XXX"
        return value
    return "XXX"

# ----------------------------
# LÓGICA PRINCIPAL
# ----------------------------
def main():
    # Verificaciones de carpeta base
    if not FOLDER_BASE.exists() or not FOLDER_BASE.is_dir():
        print(f"❌ La carpeta base no existe o no es un directorio: {FOLDER_BASE}")
        sys.exit(1)

    # Cargar Excel
    excel_path = pick_excel_path()
    print(f"📄 Usando Excel: {excel_path}")
    mapping_df = load_mapping(excel_path, SHEET_NAME, COL_SCTASK, COL_NUM)

    # Recorremos subcarpetas
    total = 0
    renombradas = 0
    saltadas = 0
    conflictos = 0

    for entry in os.listdir(FOLDER_BASE):
        old_path = FOLDER_BASE / entry
        if not old_path.is_dir():
            continue

        total += 1

        # Esperamos nombres con estructura separada por guiones
        # Ejemplo: "<NUM>-<algo>-<algo>-SCTASK1234567-..."
        fragments = entry.split("-")

        # Necesitamos al menos 4 fragmentos para extraer SCTASK del índice 3
        if len(fragments) <= 3:
            print(f"↪ Saltado (estructura insuficiente): {entry}")
            saltadas += 1
            continue

        sctask = fragments[3].strip().upper()
        if not sctask.startswith("SCTASK"):
            print(f"↪ Saltado (fragmento[3] no es SCTASK*): {entry}")
            saltadas += 1
            continue

        # Buscar NUM correspondiente
        num_value = find_num(mapping_df, COL_SCTASK, COL_NUM, sctask)
        if num_value == "XXX":
            print(f"❓ No se encontró NUM para {sctask} → se mantiene 'XXX': {entry}")

        # Construir nuevo nombre reemplazando el PRIMER fragmento por el NUM
        new_fragments = fragments[:]
        new_fragments[0] = num_value
        new_name = "-".join(new_fragments)
        new_path = FOLDER_BASE / new_name

        if new_path.exists():
            print(f"⚠ No renombrada (ya existe): {new_name}")
            conflictos += 1
            continue

        try:
            os.rename(old_path, new_path)
            print(f"✔ Renombrada: {entry}  →  {new_name}")
            renombradas += 1
        except PermissionError:
            print(f"❌ Permiso denegado al renombrar: {entry}")
        except OSError as e:
            print(f"❌ Error del sistema al renombrar '{entry}': {e}")

    print("\n--- Resumen ---")
    print(f"Total carpetas revisadas: {total}")
    print(f"Renombradas: {renombradas}")
    print(f"Saltadas: {saltadas}")
    print(f"Conflictos (destino existía): {conflictos}")

if __name__ == "__main__":
    main()
