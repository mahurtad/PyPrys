#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Validador de permisos Canvas para un curso específico.
- Marca OK si el endpoint responde 200.
- Muestra 401/403 (u otro) si no hay acceso.
- No pagina ni descarga archivos; solo HEAD/GET mínimo.

Requisitos:
  pip install requests python-dotenv  # (dotenv es opcional)

Configura:
  - CANVAS_BASE_URL (p.ej. https://cientificavirtual.cientifica.edu.pe)
  - CANVAS_TOKEN    (token de acceso personal con scopes necesarios)
  - COURSE_ID       (59178 por defecto)

Puedes definirlos como variables de entorno o editar DEFAULTS abajo.
"""

import os
import sys
import time
import requests
from urllib.parse import urlencode

# ---------- DEFAULTS (puedes editar) ----------
DEFAULT_BASE_URL = os.getenv("CANVAS_BASE_URL", "https://cientificavirtual.cientifica.edu.pe")
DEFAULT_TOKEN    = os.getenv("CANVAS_TOKEN",    "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w")
DEFAULT_COURSE   = os.getenv("COURSE_ID",       "59178")
TIMEOUT_SECONDS  = 20
# ---------------------------------------------

def build_headers(token: str):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

def check_get(session: requests.Session, url: str) -> (bool, int, str):
    try:
        r = session.get(url, timeout=TIMEOUT_SECONDS)
        # Algunos endpoints devuelven 200 pero sin datos si no hay elementos; eso sigue siendo acceso válido.
        if r.status_code == 200:
            return True, r.status_code, "OK"
        else:
            # mensaje breve
            msg = ""
            try:
                js = r.json()
                msg = js.get("errors") or js.get("message") or ""
            except Exception:
                msg = (r.text or "")[:200]
            return False, r.status_code, str(msg)[:180]
    except requests.exceptions.RequestException as e:
        return False, -1, f"EXC: {e.__class__.__name__}: {e}"

def main():
    base = (sys.argv[1] if len(sys.argv) > 1 else DEFAULT_BASE_URL).rstrip("/")
    token = (sys.argv[2] if len(sys.argv) > 2 else DEFAULT_TOKEN)
    course_id = (sys.argv[3] if len(sys.argv) > 3 else DEFAULT_COURSE)

    if not token:
        print("⚠️  Faltan credenciales. Exporta CANVAS_TOKEN o pásalo como 2do argumento.")
        print("Uso: python validate_canvas_access.py https://TU-DOMINIO token_opcional 59178_opcional")
        sys.exit(1)

    api = f"{base}/api/v1"
    s = requests.Session()
    s.headers.update(build_headers(token))

    # --------- ENDPOINTS A VALIDAR ----------
    ctx_code = f"course_{course_id}"
    endpoints = [
        # Estructura y materiales
        ("Estructura: course (syllabus_body)",
         f"{api}/courses/{course_id}?include[]=syllabus_body"),
        ("Módulos",
         f"{api}/courses/{course_id}/modules?per_page=1"),
        ("Módulos → items",
         f"{api}/courses/{course_id}/modules?include[]=items&per_page=1"),
        ("Archivos",
         f"{api}/courses/{course_id}/files?per_page=1"),
        ("Páginas",
         f"{api}/courses/{course_id}/pages?per_page=1"),
        ("Anuncios / Foros (discussion_topics)",
         f"{api}/courses/{course_id}/discussion_topics?per_page=1"),

        # Evaluaciones y fechas
        ("Tareas (assignments)",
         f"{api}/courses/{course_id}/assignments?per_page=1"),
        ("Quizzes",
         f"{api}/courses/{course_id}/quizzes?per_page=1"),
        ("Rúbricas",
         f"{api}/courses/{course_id}/rubrics?per_page=1"),

        # Calificaciones (depende del rol)
        ("Gradebook history (docente/TA)",
         f"{api}/courses/{course_id}/gradebook_history?per_page=1"),
        ("Submissions (docente/TA - todos)",
         f"{api}/courses/{course_id}/students/submissions?per_page=1"),
        ("Enrollments (docente/TA)",
         f"{api}/courses/{course_id}/enrollments?per_page=1"),
        ("Mis envíos/notas (alumno - self)",
         f"{api}/courses/{course_id}/students/submissions?student_ids[]=self&per_page=1"),

        # Calendario del curso
        ("Eventos de calendario del curso",
         f"{api}/calendar_events?{urlencode({'context_codes[]': ctx_code, 'per_page': 1})}"),
    ]
    # ----------------------------------------

    # Salida tabular mínima
    col1, col2, col3 = "Prueba", "Estado", "Detalle"
    ancho1 = max(len(col1), max(len(name) for name, _ in endpoints))
    print(f"{col1:<{ancho1}}  {col2:^8}  {col3}")
    print("-" * (ancho1 + 2 + 8 + 2 + 50))

    for name, url in endpoints:
        ok, code, msg = check_get(s, url)
        estado = "OK" if ok else (str(code) if code != -1 else "ERR")
        detalle = "OK" if ok else (msg or "—")
        print(f"{name:<{ancho1}}  {estado:^8}  {detalle[:140]}")
        # Respetar un poco para no gatillar rate limit
        time.sleep(0.2)

    print("\nLeyenda:")
    print("  OK  -> acceso válido (200).")
    print("  401 -> token inválido o expirado.")
    print("  403 -> sin permisos/rol para ese recurso.")
    print("  404 -> curso o recurso no encontrado (o sin visibilidad).")

if __name__ == "__main__":
    main()
