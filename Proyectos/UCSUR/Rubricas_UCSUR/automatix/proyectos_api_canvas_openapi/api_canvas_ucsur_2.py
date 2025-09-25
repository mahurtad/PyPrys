import requests
from collections import Counter

# ---- CONFIGURA AQUÍ ----
BASE_URL = "https://cientificavirtual.cientifica.edu.pe/api/v1"
ACCOUNT_ID = 2070 
TOKEN = "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w"           # <-- pega tu token
TERM_NAME_FILTER = None             # ej. "Pregrado 2252" o deja None para no filtrar
PER_PAGE = 100
# ------------------------

session = requests.Session()
session.headers.update({"Authorization": f"Bearer {TOKEN}"})


def next_link(headers):
    """Devuelve la URL 'next' del header Link de Canvas (si existe)."""
    link = headers.get("Link")
    if not link:
        return None
    for part in link.split(","):
        url, rel = part.split(";")
        if 'rel="next"' in rel:
            return url.strip().lstrip("<").rstrip(">")
    return None


def get_enrollment_term_id_by_name(account_id, term_name):
    """Busca un término por nombre y retorna su ID (o None si no lo encuentra)."""
    url = f"{BASE_URL}/accounts/{account_id}/terms"
    r = session.get(url)
    r.raise_for_status()
    for t in r.json().get("enrollment_terms", []):
        if t.get("name") == term_name:
            return t.get("id")
    return None


def get_all_courses(account_id, term_id=None):
    """Trae todos los cursos de la cuenta (y subcuentas), con paginación."""
    params = {"with_subaccounts": "true", "per_page": PER_PAGE}
    if term_id:
        params["enrollment_term_id"] = term_id

    url = f"{BASE_URL}/accounts/{account_id}/courses"
    data = []
    while url:
        r = session.get(url, params=params if "?" not in url else None)
        if r.status_code != 200:
            raise SystemExit(f"Error {r.status_code}: {r.text}")
        data.extend(r.json())
        url = next_link(r.headers)
        params = None
    return data


def get_account_names(account_ids):
    """Devuelve dict {account_id: account_name} consultando /accounts/:id."""
    names = {}
    for aid in account_ids:
        r = session.get(f"{BASE_URL}/accounts/{aid}")
        if r.status_code == 200:
            names[aid] = r.json().get("name") or str(aid)
        else:
            names[aid] = str(aid)
    return names


def main():
    term_id = None
    if TERM_NAME_FILTER:
        term_id = get_enrollment_term_id_by_name(ACCOUNT_ID, TERM_NAME_FILTER)
        if not term_id:
            print(f"⚠️ No se encontró el término '{TERM_NAME_FILTER}'. Se listarán todos los cursos.")
    courses = get_all_courses(ACCOUNT_ID, term_id)

    # Agrupar por subcuenta
    by_account_id = Counter(c.get("account_id") for c in courses if c.get("account_id"))
    account_names = get_account_names(set(by_account_id.keys()))

    total_cursos = sum(by_account_id.values())
    total_carreras = len(by_account_id)

    # Mostrar tabla
    print("\n=== Cursos por subcuenta (carrera) ===")
    print(f"{'Carrera/Subcuenta':50} | {'# Cursos':>8}")
    print("-" * 65)
    for aid, count in sorted(by_account_id.items(), key=lambda x: x[1], reverse=True):
        name = account_names.get(aid, str(aid))
        print(f"{name:50} | {count:8}")

    print("-" * 65)
    print(f"{'TOTAL':50} | {total_cursos:8}")
    print(f"Cantidad de carreras distintas: {total_carreras}")


if __name__ == "__main__":
    main()
