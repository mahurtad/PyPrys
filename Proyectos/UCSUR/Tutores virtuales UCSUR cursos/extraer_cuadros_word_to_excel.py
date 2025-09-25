import re
import sys
import traceback
from pathlib import Path
from typing import List, Tuple, Optional

import pandas as pd

# ---- Configuración ----
BASE_DIR = r"C:\Users\manue\Downloads\Silabos\Version final Silabos 4"
ENCODING_FALLBACK = "utf-8"
MAKE_LOG = True
WRITE_EMPTY_6C = True  # Si no hay columnas de actividades, crea 6 columnas con vacío

# Inferencia de MÓDULO según SEMANA (si falta MÓDULO o está vacío):
# Módulo 1 = semanas 1-3, Módulo 2 = semanas 4-7, Módulo 3 = semanas 8-10, Módulo 4 = semanas 11-16
INFER_MODULE_FROM_WEEK = True

# Columnas objetivo
COL_4 = ["MÓDULO DE APRENDIZAJE", "SEMANA", "SESIÓN", "TEMAS"]
COL_6 = COL_4 + ["ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE", "ACTIVIDADES DE TRABAJO AUTÓNOMO"]

# Mapeo de sinónimos -> nombre estándar
HEADER_ALIASES = {
    "modulo de aprendizaje": "MÓDULO DE APRENDIZAJE",
    "módulo de aprendizaje": "MÓDULO DE APRENDIZAJE",
    "modulo": "MÓDULO DE APRENDIZAJE",
    "módulo": "MÓDULO DE APRENDIZAJE",
    "semana": "SEMANA",
    "sesion": "SESIÓN",
    "sesión": "SESIÓN",
    "temas": "TEMAS",
    "actividad docente": "ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE",
    "actividades en interacción con el docente": "ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE",
    "actividades en interaccion con el docente": "ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE",
    "actividad de trabajo autónomo": "ACTIVIDADES DE TRABAJO AUTÓNOMO",
    "actividades de trabajo autónomo": "ACTIVIDADES DE TRABAJO AUTÓNOMO",
    "actividades de trabajo autonomo": "ACTIVIDADES DE TRABAJO AUTÓNOMO",
    "trabajo autonomo": "ACTIVIDADES DE TRABAJO AUTÓNOMO",
}

# Señaladores de sección 8
SECTION_8_PATTERNS = [
    r"\b8\.\s*actividades\s*principales\b",
    r"\bsecci[oó]n\s*8\b.*actividades",
]

# ---- Intentar importar parsers opcionales ----
_has_camelot = False
try:
    import camelot
    _has_camelot = True
except Exception:
    _has_camelot = False

_has_tabula = False
try:
    import tabula  # noqa
    _has_tabula = True
except Exception:
    _has_tabula = False

import pdfplumber
from docx import Document


# ----------------- Utilidades -----------------
def normalize(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_key(s: str) -> str:
    return normalize(s).lower().replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")

def map_headers(cols: List[str]) -> List[str]:
    mapped = []
    for c in cols:
        key = normalize_key(c)
        std = HEADER_ALIASES.get(key)
        if not std:
            # aproximación por inclusión
            if "mod" in key:
                std = "MÓDULO DE APRENDIZAJE"
            elif "sem" in key:
                std = "SEMANA"
            elif "ses" in key:
                std = "SESIÓN"
            elif "tema" in key:
                std = "TEMAS"
            elif "docent" in key or ("interacc" in key and "doc" in key):
                std = "ACTIVIDADES EN INTERACCIÓN CON EL DOCENTE"
            elif "autonom" in key or "autónom" in key or "trabajo" in key:
                std = "ACTIVIDADES DE TRABAJO AUTÓNOMO"
            else:
                std = c  # deja como está
        mapped.append(std)
    return mapped

def coerce_int(x):
    try:
        xi = int(str(x).strip())
        return xi
    except Exception:
        return x  # deja como texto si no se puede

def infer_module_from_week(week_val) -> Optional[str]:
    try:
        w = int(str(week_val).strip())
    except Exception:
        return None
    if 1 <= w <= 3:
        return "1"
    if 4 <= w <= 7:
        return "2"
    if 8 <= w <= 10:
        return "3"
    if 11 <= w <= 16:
        return "4"
    return None

def postprocess_df(df: pd.DataFrame) -> pd.DataFrame:
    # Normalizar encabezados
    df.columns = map_headers(list(df.columns))

    # Mantener solo columnas relevantes (si existen)
    # Primero, intenta identificar columnas que contienen las 4 o 6 requeridas
    keep_cols = [c for c in COL_6 if c in df.columns]
    if not keep_cols:
        keep_cols = [c for c in COL_4 if c in df.columns]

    df = df[keep_cols].copy()
    # limpiar espacios
    for c in df.columns:
        df[c] = df[c].apply(normalize)

    # Limpiar filas totalmente vacías
    df = df.replace({"None": "", "nan": "", "NaN": ""})
    df = df[~(df.fillna("").apply(lambda r: all(len(str(v).strip()) == 0 for v in r), axis=1))]

    # Coaccionar numéricos para SEMANA/SESIÓN cuando se pueda
    if "SEMANA" in df.columns:
        df["SEMANA"] = df["SEMANA"].apply(coerce_int)
    if "SESIÓN" in df.columns:
        df["SESIÓN"] = df["SESIÓN"].apply(coerce_int)

    # Si falta MÓDULO o está vacío, infiere desde SEMANA
    if "MÓDULO DE APRENDIZAJE" not in df.columns:
        df["MÓDULO DE APRENDIZAJE"] = df.get("SEMANA", pd.Series([None]*len(df))).apply(infer_module_from_week)
    else:
        # completar vacíos si procede
        if INFER_MODULE_FROM_WEEK:
            df["MÓDULO DE APRENDIZAJE"] = df.apply(
                lambda r: r["MÓDULO DE APRENDIZAJE"] if str(r["MÓDULO DE APRENDIZAJE"]).strip() not in ("", "None", "nan")
                else infer_module_from_week(r.get("SEMANA")), axis=1
            )

    # Dejar solo el número en MÓDULO (si viene como "Módulo 1", etc.)
    if "MÓDULO DE APRENDIZAJE" in df.columns:
        df["MÓDULO DE APRENDIZAJE"] = df["MÓDULO DE APRENDIZAJE"].apply(
            lambda x: re.sub(r"(?i)m[oó]dulo\s*", "", str(x)).strip() if x is not None else x
        )

    # Asegura columnas target en orden deseado (llenando faltantes con vacío)
    for col in COL_6:
        if col not in df.columns:
            df[col] = ""

    df = df[COL_6]  # orden canon
    return df

def find_section8_pages_pdf(pdf_path: Path) -> List[int]:
    pages = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            txt = page.extract_text() or ""
            tnorm = normalize_key(txt)
            if any(re.search(pat, tnorm) for pat in SECTION_8_PATTERNS):
                pages.append(i)  # índice base 0
    # si no encontró por texto, devuelve todas (último recurso)
    return pages or list(range(0, 9999))

def extract_tables_pdf(pdf_path: Path) -> List[pd.DataFrame]:
    """
    Intenta por orden:
    - Camelot (lattice y stream)
    - pdfplumber (fallback)
    Filtra a páginas donde aparece sección 8 si es posible.
    """
    dfs: List[pd.DataFrame] = []
    target_pages = find_section8_pages_pdf(pdf_path)
    # recorta a ventanas de 1-3 páginas siguientes por seguridad
    window_pages = []
    for p in target_pages:
        window_pages.extend([p, p+1, p+2])
    window_pages = sorted(set([p for p in window_pages if p >= 0]))

    # 1) Camelot
    if _has_camelot:
        try:
            pages_arg = ",".join(str(p+1) for p in window_pages) if window_pages else "all"
            for flavor in ("lattice", "stream"):
                try:
                    tables = camelot.read_pdf(str(pdf_path), pages=pages_arg, flavor=flavor)
                    for t in tables:
                        df = t.df.copy()
                        # primera fila como encabezado si parece encabezado
                        if df.shape[0] > 1:
                            df.columns = df.iloc[0].tolist()
                            df = df.iloc[1:].reset_index(drop=True)
                        dfs.append(df)
                except Exception:
                    pass
        except Exception:
            pass

    # 2) pdfplumber fallback
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            last_page_index = len(pdf.pages) - 1
            for p in window_pages or range(len(pdf.pages)):
                if p > last_page_index:
                    continue
                page = pdf.pages[p]
                # intenta varias configuraciones
                for ts in [
                    {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
                    {"vertical_strategy": "text", "horizontal_strategy": "text"},
                    {"vertical_strategy": "lines_strict", "horizontal_strategy": "lines_strict"},
                ]:
                    try:
                        tbls = page.extract_tables(table_settings=ts)
                        for raw in tbls or []:
                            df = pd.DataFrame(raw)
                            # usa primera fila como header si parece encabezado
                            if df.shape[0] > 1:
                                df.columns = df.iloc[0].tolist()
                                df = df.iloc[1:].reset_index(drop=True)
                            dfs.append(df)
                    except Exception:
                        continue
    except Exception:
        pass

    # Limpieza: quitar dfs vacíos/ruido evidente
    clean = []
    for d in dfs:
        if d is None or d.empty:
            continue
        # descartar tablas con 1 columna solamente sin contenido
        if d.shape[1] <= 1:
            continue
        clean.append(d)
    return clean

def extract_tables_docx(docx_path: Path) -> List[pd.DataFrame]:
    """
    Busca la sección 8 por texto en párrafos y a partir de ahí captura tablas
    hasta que aparezca la sección 9 (o fin).
    """
    doc = Document(str(docx_path))
    # localizar índice de párrafo de sección 8
    sec8_idx = None
    for i, p in enumerate(doc.paragraphs):
        t = normalize_key(p.text)
        if any(re.search(pat, t) for pat in SECTION_8_PATTERNS):
            sec8_idx = i
            break

    # recolectar tablas desde sec8_idx en adelante hasta hallar "9."
    tables = []
    start_collect = sec8_idx is None  # si no encontró, toma todas (mejor algo que nada)
    for block in doc.element.body.iter():
        tag = block.tag.lower()
        # saltar hasta encontrar sec8
        if not start_collect and "p" in tag:
            # texto del párrafo
            try:
                txt = normalize_key(block.text)
                if any(re.search(pat, txt) for pat in SECTION_8_PATTERNS):
                    start_collect = True
                    continue
            except Exception:
                pass

        if start_collect and "tbl" in tag:
            # convertir tabla a df
            rows = []
            for row in block.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr"):
                cells = []
                for cell in row.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc"):
                    # concatenar párrafos
                    txt = " ".join(
                        normalize("".join([r.text or "" for r in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")]))
                        for p in cell.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p")
                    )
                    cells.append(txt)
                rows.append(cells)
            if rows:
                df = pd.DataFrame(rows)
                if df.shape[0] > 1:
                    df.columns = df.iloc[0].tolist()
                    df = df.iloc[1:].reset_index(drop=True)
                tables.append(df)

        # detener si aparece sección 9
        if start_collect and "p" in tag:
            try:
                txt = normalize_key(block.text)
                if re.search(r"\b9\.\b", txt):
                    break
            except Exception:
                pass

    return tables

def stitch_and_clean(dfs: List[pd.DataFrame]) -> pd.DataFrame:
    """
    Une varias tablas potencialmente fragmentadas, normaliza encabezados
    y devuelve un único DataFrame postprocesado.
    """
    if not dfs:
        return pd.DataFrame(columns=COL_6)
    # normaliza cada df y quédate con columnas de interés
    cleaned = []
    for d in dfs:
        try:
            d2 = postprocess_df(d)
            cleaned.append(d2)
        except Exception:
            continue
    if not cleaned:
        return pd.DataFrame(columns=COL_6)
    big = pd.concat(cleaned, ignore_index=True)
    # quitar duplicados obvios
    big = big.drop_duplicates()
    # re-postprocess para asegurar orden/columnas completas
    big = postprocess_df(big)
    return big

def process_file(path: Path) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], str]:
    """
    Devuelve (df4, df6, mensaje_log)
    """
    ext = path.suffix.lower()
    dfs: List[pd.DataFrame] = []
    try:
        if ext == ".pdf":
            dfs = extract_tables_pdf(path)
        elif ext in (".docx",):
            dfs = extract_tables_docx(path)
        else:
            return None, None, f"SKIP: extensión no soportada: {ext}"

        if not dfs:
            return None, None, "SIN TABLAS DETECTADAS EN SECCIÓN 8"

        df6 = stitch_and_clean(dfs)
        # derivar df4
        df4 = df6[COL_4].copy()

        # Si no había actividades y WRITE_EMPTY_6C=True, de todas formas entrega 6c
        if not any(df6.get(c, "").astype(str).str.strip().any() for c in COL_6[4:]) and WRITE_EMPTY_6C:
            # ya están vacías por postprocess, se entregan así
            pass

        return df4, df6, f"OK: filas={len(df6)}"
    except Exception as e:
        return None, None, f"ERROR: {e}"

def write_outputs(path: Path, df4: Optional[pd.DataFrame], df6: Optional[pd.DataFrame]) -> List[Path]:
    outs = []
    if df4 is not None and not df4.empty:
        out4 = path.with_name(f"Plan_4Columnas_{path.stem}.xlsx")
        df4.to_excel(out4, index=False)
        outs.append(out4)
    if df6 is not None and not df6.empty:
        out6 = path.with_name(f"Plan_6Columnas_{path.stem}.xlsx")
        df6.to_excel(out6, index=False)
        outs.append(out6)
    return outs

def main():
    base = Path(BASE_DIR)
    if not base.exists():
        print(f"[ERROR] No existe la ruta base: {base}")
        sys.exit(1)

    files = []
    for ext in ("*.pdf", "*.docx"):
        files.extend(base.rglob(ext))

    if not files:
        print("[AVISO] No se encontraron archivos PDF/DOCX.")
        sys.exit(0)

    log_rows = []
    ok, fail, skip = 0, 0, 0

    for f in sorted(files):
        print(f"→ Procesando: {f}")
        df4, df6, msg = process_file(f)
        try:
            outs = write_outputs(f, df4, df6)
            if outs:
                print("   ✓ Generados:", ", ".join(str(o) for o in outs))
                ok += 1
                log_rows.append({"archivo": str(f), "resultado": msg, "salidas": "; ".join(str(o) for o in outs)})
            else:
                # si no hubo salidas pero no es error de extensión
                if msg.startswith("SKIP"):
                    skip += 1
                else:
                    fail += 1
                print("   ✗", msg)
                log_rows.append({"archivo": str(f), "resultado": msg, "salidas": ""})
        except Exception as e:
            fail += 1
            emsg = f"ERROR al escribir: {e}"
            print("   ✗", emsg)
            log_rows.append({"archivo": str(f), "resultado": emsg, "salidas": ""})

    print(f"\n[RESUMEN] OK: {ok} | Fallos: {fail} | Omitidos: {skip}")

    if MAKE_LOG:
        log_path = base / "extract_log.csv"
        pd.DataFrame(log_rows).to_csv(log_path, index=False, encoding="utf-8-sig")
        print(f"[LOG] {log_path}")

if __name__ == "__main__":
    main()
