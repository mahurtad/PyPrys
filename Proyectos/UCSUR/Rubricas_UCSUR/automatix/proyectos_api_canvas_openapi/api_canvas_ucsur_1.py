import requests

BASE_URL = "https://cientificavirtual.cientifica.edu.pe/api/v1"
TOKEN = "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w"  # <-- pega tu token Canvas
ACCOUNT_ID = 2070

session = requests.Session()
session.headers.update({"Authorization": f"Bearer {TOKEN}"})

def get_all_courses(account_id, search_term=None, enrollment_term_id=None,
                    with_subaccounts=True, published=None, per_page=100):
    """Itera todas las páginas y devuelve la lista completa de cursos."""
    params = {
        "with_subaccounts": str(with_subaccounts).lower(),
        "per_page": per_page,
    }
    if search_term:
        params["search_term"] = search_term
    if enrollment_term_id:
        params["enrollment_term_id"] = enrollment_term_id
    if published is not None:
        params["published"] = str(bool(published)).lower()

    url = f"{BASE_URL}/accounts/{account_id}/courses"
    courses = []

    while url:
        r = session.get(url, params=params if '?' not in url else None)
        if r.status_code != 200:
            raise SystemExit(f"Error {r.status_code}: {r.text}")
        courses.extend(r.json())
        # paginación por encabezado Link
        next_url = None
        if "Link" in r.headers:
            for part in r.headers["Link"].split(","):
                segs = part.split(";")
                if len(segs) >= 2 and 'rel="next"' in segs[1]:
                    next_url = segs[0].strip().lstrip("<").rstrip(">")
                    break
        url = next_url
        params = None  # después de la primera llamada, Canvas ya incluye querys en next

    return courses

# --- Ejecución ---
cursos = get_all_courses(ACCOUNT_ID, with_subaccounts=True)
print(f"Total cursos: {len(cursos)}\n")

for c in cursos[:30]:  # muestra los primeros 30
    print(f"- {c.get('id')} | {c.get('name')} "
          f"| SIS: {c.get('sis_course_id')} "
          f"| Term: {c.get('enrollment_term_id')} "
          f"| Estado: {c.get('workflow_state')}")

# (Opcional) Guardar a CSV
import csv
with open("cursos_account_2070.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["id","name","course_code","sis_course_id","enrollment_term_id","workflow_state","account_id","uuid"])
    for c in cursos:
        w.writerow([
            c.get("id"), c.get("name"), c.get("course_code"),
            c.get("sis_course_id"), c.get("enrollment_term_id"),
            c.get("workflow_state"), c.get("account_id"), c.get("uuid")
        ])
print("\nArchivo generado: cursos_account_2070.csv")
