import requests, csv

BASE_URL = "https://cientificavirtual.cientifica.edu.pe/api/v1"
TOKEN = "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w"   # usa el mismo que tienes en tu código
ACCOUNT_ID = 2070

session = requests.Session()
session.headers.update({"Authorization": f"Bearer {TOKEN}"})


def get_all_courses(account_id, per_page=100):
    """Obtiene todos los cursos del account."""
    url = f"{BASE_URL}/accounts/{account_id}/courses"
    params = {"with_subaccounts": "true", "per_page": per_page}
    cursos = []

    while url:
        r = session.get(url, params=params if '?' not in url else None)
        r.raise_for_status()
        cursos.extend(r.json())
        # Paginación
        next_url = None
        if "Link" in r.headers:
            for part in r.headers["Link"].split(","):
                if 'rel="next"' in part:
                    next_url = part.split(";")[0].strip().strip("<>")
                    break
        url = next_url
        params = None
    return cursos


def get_student_count(course_id):
    """Devuelve el número de alumnos matriculados en un curso."""
    url = f"{BASE_URL}/courses/{course_id}/enrollments"
    params = {"type[]": "StudentEnrollment", "per_page": 100}
    count = 0
    while url:
        r = session.get(url, params=params if '?' not in url else None)
        r.raise_for_status()
        data = r.json()
        count += len(data)
        # Paginación
        next_url = None
        if "Link" in r.headers:
            for part in r.headers["Link"].split(","):
                if 'rel="next"' in part:
                    next_url = part.split(";")[0].strip().strip("<>")
                    break
        url = next_url
        params = None
    return count


# --- Ejecución ---
cursos = get_all_courses(ACCOUNT_ID)
print(f"Total cursos: {len(cursos)}\n")

tabla = []
for c in cursos:
    cid = c.get("id")
    name = c.get("name")
    alumnos = get_student_count(cid)
    tabla.append([cid, name, alumnos])
    print(f"- {cid} | {name} | Alumnos: {alumnos}")

# Exportar a CSV
with open("alumnos_por_curso.csv", "w", newline="", encoding="utf-8") as f:
    w = csv.writer(f)
    w.writerow(["course_id", "course_name", "num_alumnos"])
    w.writerows(tabla)

print("\nArchivo generado: alumnos_por_curso.csv")
