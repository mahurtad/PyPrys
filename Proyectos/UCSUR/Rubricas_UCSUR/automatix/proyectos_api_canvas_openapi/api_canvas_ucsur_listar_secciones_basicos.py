#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, re, time, requests
import pandas as pd
from typing import Dict, Iterable, List, Tuple

# -------------------- Config --------------------
BASE_URL   = (sys.argv[1] if len(sys.argv) > 1 else os.getenv("CANVAS_BASE_URL", "https://cientificavirtual.cientifica.edu.pe")).rstrip("/")
TOKEN      = (sys.argv[2] if len(sys.argv) > 2 else os.getenv("CANVAS_TOKEN", "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w"))
ACCOUNT_ID = (sys.argv[3] if len(sys.argv) > 3 else os.getenv("CANVAS_ACCOUNT_ID", "2124"))
NAME_FILTER= (sys.argv[4] if len(sys.argv) > 4 else os.getenv("NAME_FILTER", "álgebra"))
OUT_XLSX   = os.getenv("OUT_XLSX", f"secciones_{re.sub(r'\\W+', '_', NAME_FILTER.lower())}_{ACCOUNT_ID}.xlsx")

TIMEOUT    = 25
SLEEP      = 0.15
# ------------------------------------------------

if not TOKEN:
    raise SystemExit("❗ Falta CANVAS_TOKEN (o pásalo como 2º argumento).")

API = f"{BASE_URL}/api/v1"
HEADERS = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}

def paginated_get(url: str, params: Dict = None) -> Iterable[Dict]:
    s = requests.Session()
    s.headers.update(HEADERS)
    next_url = url
    while next_url:
        r = s.get(next_url, params=params, timeout=TIMEOUT)
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list):
            for item in data:
                yield item
        else:
            for item in data.get("courses", []):
                yield item
        next_url = None
        link = r.headers.get("Link", "")
        for part in link.split(","):
            if 'rel="next"' in part:
                next_url = part[part.find("<")+1:part.find(">")]
                break
        time.sleep(SLEEP)

def list_courses_in_account(account_id: str) -> Iterable[Dict]:
    return paginated_get(f"{API}/accounts/{account_id}/courses", {"per_page": 100})

def iter_sections(course_id: int) -> Iterable[Dict]:
    return paginated_get(f"{API}/courses/{course_id}/sections", {"per_page": 100})

def count_active_students_in_section(section_id: int) -> int:
    url = f"{API}/sections/{section_id}/enrollments"
    params = {"per_page": 100, "type[]": "StudentEnrollment", "state[]": "active"}
    return sum(1 for _ in paginated_get(url, params))

def normalize(s: str) -> str:
    return (s.lower()
            .replace("á","a").replace("é","e").replace("í","i")
            .replace("ó","o").replace("ú","u").replace("ñ","n"))

def main():
    name_pat = normalize(NAME_FILTER)
    rows: List[Tuple[int, str, int, str, int]] = []

    for c in list_courses_in_account(ACCOUNT_ID):
        course_id = c.get("id")
        course_name = c.get("name") or c.get("course_code") or ""
        if not course_id or not course_name:
            continue
        if name_pat in normalize(course_name):
            for sec in iter_sections(course_id):
                sid = sec.get("id")
                sname = sec.get("name") or ""
                if not sid:
                    continue
                try:
                    students = count_active_students_in_section(sid)
                except Exception:
                    students = -1  # señal de error/permiso
                rows.append((course_id, course_name, sid, sname, students))

    # ---- Exportar a Excel ----
    df = pd.DataFrame(rows, columns=[
        "course_id", "course_name", "section_id", "section_name", "students_count"
    ])

    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Secciones")

        # Autoajustar ancho simple
        ws = xw.sheets["Secciones"]
        for col_idx, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).values])
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 60)

    print(f"✅ Exportado {len(df)} filas a: {OUT_XLSX}")
    if (df["students_count"] == -1).any():
        print("⚠️ Algunas secciones no pudieron contarse (marcadas como -1). Verifica permisos del token.")

if __name__ == "__main__":
    main()
