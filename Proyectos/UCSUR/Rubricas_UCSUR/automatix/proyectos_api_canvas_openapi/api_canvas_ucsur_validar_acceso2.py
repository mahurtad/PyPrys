#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Validador de permisos Canvas para un curso específico (extendido).
- Marca OK si el endpoint responde 200.
- Muestra 401/403/404 (u otro) si no hay acceso.
- Incluye validaciones de Page Views (admin/self) y Analytics.

Requisitos:
  pip install requests python-dotenv

Configura:
  - CANVAS_BASE_URL (p.ej. https://cientificavirtual.cientifica.edu.pe)
  - CANVAS_TOKEN    (token personal con scopes necesarios)
  - COURSE_ID       (p.ej. 59178)
Puedes definirlos como variables de entorno o pasarlos como argumentos.
"""

import os
import sys
import time
import requests
from urllib.parse import urlencode

# ---------- DEFAULTS (sin token hardcodeado) ----------
DEFAULT_BASE_URL = os.getenv("CANVAS_BASE_URL", "https://cientificavirtual.cientifica.edu.pe")
DEFAULT_TOKEN    = os.getenv("CANVAS_TOKEN", "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w")
DEFAULT_COURSE   = os.getenv("COURSE_ID", "59178")
TIMEOUT_SECONDS  = 20
# ------------------------------------------------------

def build_headers(token: str):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

def check_get(session: requests.Session, url: str) -> (bool, int, str):
    try:
        r = session.get(url, timeout=TIMEOUT_SECONDS)
        if r.status_code == 200:
            return True, r.status_code, "OK"
        else:
            msg = ""
            try:
                js = r.json()
                msg = js.get("errors") or js.get("message") or js
            except Exception:
                msg = (r.text or "")[:200]
            return False, r.status_code, str(msg)[:180]
    except requests.exceptions.RequestException as e:
        return False, -1, f"EXC: {e.__class__.__name__}: {e}"

def get_json(session: requests.Session, url: str):
    ok, code, _ = check_get(session, url)
    if not ok:
        return None, code
    try:
        r = session.get(url, timeout=TIMEOUT_SECONDS)
        if r.status_code == 200:
            return r.json(), 200
        return None, r.status_code
    except requests.exceptions.RequestException:
        return None, -1

def find_sample_student_id(session: requests.Session, api: str, course_id: str):
    """Devuelve un student_id del curso (si existe) para probar page_views/analytics."""
    url = f"{api}/courses/{course_id}/users?enrollment_type[]=student&per_page=1"
    data, code = get_json(session, url)
    if code == 200 and isinstance(data, list) and data:
        return data[0].get("id")
    return None

def main():
    base = (sys.argv[1] if len(sys.argv) > 1 else DEFAULT_BASE_URL).rstrip("/")
    token = (sys.argv[2] if len(sys.argv) > 2 else DEFAULT_TOKEN)
    course_id = (sys.argv[3] if len(sys.argv) > 3 else DEFAULT_COURSE)

    if not token:
        print("⚠️  Faltan credenciales. Exporta CANVAS_TOKEN o pásalo como 2do argumento.")
        print("Uso: python validar_canvas.py https://TU-DOMINIO token 59178")
        sys.exit(1)

    api = f"{base}/api/v1"
    s = requests.Session()
    s.headers.update(build_headers(token))

    # --------- obtener un alumno de prueba (si hay) ----------
    sample_student_id = find_sample_student_id(s, api, course_id)
    # --------------------------------------------------------

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

        # Roster para pruebas
        ("Users (1 alumno para pruebas)",
         f"{api}/courses/{course_id}/users?enrollment_type[]=student&per_page=1"),
    ]

    # Añadimos dinámicamente pruebas que dependen de student_id
    if sample_student_id:
        endpoints.extend([
            # Page Views (admin)
            (f"Page Views (admin) alumno {sample_student_id}",
             f"{api}/users/{sample_student_id}/page_views?per_page=1"),
            # Analytics (resumen y detalle)
            ("Analytics: student_summaries",
             f"{api}/courses/{course_id}/analytics/student_summaries?per_page=1"),
            (f"Analytics: actividad alumno {sample_student_id}",
             f"{api}/courses/{course_id}/analytics/users/{sample_student_id}/activity"),
            (f"Analytics: assignments alumno {sample_student_id}",
             f"{api}/courses/{course_id}/analytics/users/{sample_student_id}/assignments"),
        ])
    else:
        # Si no hay alumno, probamos page_views del propio token (self)
        endpoints.append(
            ("Page Views (self)", f"{api}/users/self/page_views?per_page=1")
        )

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
        time.sleep(0.2)  # evita rate limit

    print("\nLeyenda:")
    print("  OK  -> acceso válido (200).")
    print("  401 -> token inválido o expirado.")
    print("  403 -> sin permisos/rol para ese recurso.")
    print("  404 -> curso o recurso no encontrado (o sin visibilidad).")
    print("\nNotas:")
    print("  • Page Views (admin) requiere permiso 'View users’ page views' a nivel cuenta.")
    print("  • Analytics puede estar deshabilitado por política institucional.")
    print("  • Para 'Submissions (todos)' el token debe ser docente/TA del curso o admin con permiso equivalente.")
    print("  • Quita/anonimiza IP/UA si vas a exponerlos a docentes (privacidad).")

if __name__ == "__main__":
    main()
