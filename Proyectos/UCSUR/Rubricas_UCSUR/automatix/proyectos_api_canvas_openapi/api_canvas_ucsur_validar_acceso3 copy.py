#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Reporte ADMIN por consola de TODOS los cursos del alumno (SIS) dentro de la SUBCUENTA 2070
(incluye también cursos en sus subcuentas), con métricas de Page Views y última entrega (Analytics).

Requisitos de permisos en la subcuenta 2070:
  - Users: view list
  - Read SIS data
  - View users' page views
  - (opcional) View course analytics
"""

import os, sys, re, datetime as dt, requests

ACCOUNT_ID = 2070  # subcuenta objetivo
BASE = os.getenv("CANVAS_BASE_URL", "https://cientificavirtual.cientifica.edu.pe").rstrip("/")
TOKEN = os.getenv("CANVAS_TOKEN", "yYZtyLknc76HVYnxnhAvC8NTZruMTuwJGWceEvk9BrrwvmehCDfkyGTT7rmMuk9w")
DEFAULT_SIS = os.getenv("SIS_CODE", "100175006")
API = f"{BASE}/api/v1"

# ---------------- utils ----------------
def session_or_die():
    if not BASE:
        print("Falta CANVAS_BASE_URL (ej.: https://cientificavirtual.cientifica.edu.pe)"); sys.exit(1)
    if not TOKEN:
        print("Falta CANVAS_TOKEN. Exporta la variable."); sys.exit(1)
    s = requests.Session()
    s.headers.update({"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"})
    return s

def fetch_all(sess, url, params=None):
    params = params or {}; params.setdefault("per_page", 100)
    out = []
    while url:
        r = sess.get(url, params=params, timeout=40)
        if r.status_code != 200:
            raise requests.HTTPError(f"{r.status_code} {r.text[:200]}")
        js = r.json()
        out.extend(js if isinstance(js, list) else [js])
        link = r.headers.get("Link", "")
        m = re.search(r'<([^>]+)>;\s*rel="next"', link)
        url = m.group(1) if m else None
        params = None
    return out

def parse_ua(ua):
    if not ua: return "Unknown","Unknown"
    b = "Edge" if "Edg" in ua else ("Firefox" if "Firefox" in ua else ("Chrome" if "Chrome" in ua and "Edg" not in ua and "Chromium" not in ua else ("Safari" if "Safari" in ua and "Chrome" not in ua else "Unknown")))
    osn = "Windows" if "Windows" in ua else ("macOS" if ("Mac OS X" in ua or "Macintosh" in ua) else ("Android" if "Android" in ua else ("iOS/iPadOS" if ("iPhone" in ua or "iPad" in ua) else ("Linux" if "Linux" in ua else "Unknown"))))
    return b, osn

def iso_human(ts):
    if not ts: return "-"
    return dt.datetime.fromisoformat(ts.replace("Z","+00:00")).strftime("%Y-%m-%d %H:%M:%S")

def hms(seconds):
    secs = int(seconds or 0); return f"{secs//3600:d}:{(secs%3600)//60:02d}:{secs%60:02d}"

# ------------- admin: usuario -------------
def get_user_id_by_sis(sess, sis_code):
    for path in (f"{API}/users/sis_user_id:{sis_code}",
                 f"{API}/users/sis_login_id:{sis_code}",
                 f"{API}/users/sis_user_id:{sis_code}/profile"):
        r = sess.get(path, timeout=20)
        if r.status_code == 200:
            js = r.json()
            uid = js.get("id") or js.get("user",{}).get("id")
            if uid: return uid, js
    return None, None

# ------------- admin: cursos de la subcuenta -------------
def list_courses_for_user_in_account(sess, account_id, user_id):
    """
    Ruta ADMIN en la subcuenta 2070. NO filtramos por state[] para incluir unpublished/created/claimed.
    """
    params = {
        "enrollment_user_id[]": user_id,
        "enrollment_type[]": "student",
        "with_enrollments": "true",
        "include[]": "subaccount",   # también devuelve info de subcuenta del curso
        "per_page": 100
    }
    url = f"{API}/accounts/{account_id}/courses"
    return fetch_all(sess, url, params=params)

def get_course_info(sess, course_id):
    r = sess.get(f"{API}/courses/{course_id}", timeout=20)
    if r.status_code == 200:
        js = r.json()
        return {"name": js.get("name") or js.get("course_code") or f"Course {course_id}",
                "account_id": js.get("account_id"),
                "root_account_id": js.get("root_account_id")}
    return {"name": f"Course {course_id}", "account_id": None, "root_account_id": None}

# cache de parent_account_id para subir el árbol
_account_parent_cache = {}
def get_account(sess, acc_id):
    if acc_id in _account_parent_cache:
        return _account_parent_cache[acc_id]
    r = sess.get(f"{API}/accounts/{acc_id}", timeout=20)
    if r.status_code == 200:
        js = r.json()
        _account_parent_cache[acc_id] = js
        return js
    _account_parent_cache[acc_id] = {}
    return {}

def is_in_account_tree(sess, account_id, target_id):
    """
    ¿account_id pertenece al subárbol de target_id (o es igual)?
    Sube por parent_account_id hasta root.
    """
    if account_id is None: return False
    cur = int(account_id); tgt = int(target_id); visited = set()
    while cur and cur not in visited:
        if cur == tgt: return True
        visited.add(cur)
        info = get_account(sess, cur)
        cur = info.get("parent_account_id")
    return False

# ------------- enrollment en curso -------------
def get_user_enrollment_in_course_admin(sess, course_id, user_id):
    users = fetch_all(sess, f"{API}/courses/{course_id}/users",
                      params={"user_id": user_id, "include[]": "enrollments", "per_page": 50})
    if users:
        u = users[0]
        enr = (u.get("enrollments") or [{}])[0]
        return {"name": u.get("name"),
                "login_id": u.get("login_id") or u.get("email"),
                "role": enr.get("type"),
                "state": enr.get("enrollment_state") or enr.get("workflow_state")}
    return {"name": None, "login_id": None, "role": None, "state": None}

# ------------- Page Views / Analytics -------------
def page_views_by_course(sess, user_id, since_iso=None, until_iso=None):
    params = {}
    if since_iso: params["start_time"] = since_iso
    if until_iso: params["end_time"] = until_iso
    pvs = fetch_all(sess, f"{API}/users/{user_id}/page_views", params=params)
    per_course = {}
    for ev in pvs:
        if ev.get("context_type") != "Course": continue
        cid = ev.get("context_id")
        g = per_course.setdefault(cid, {"time_spent":0,"first":None,"last":None,"ua":"","ip":"","last_type":"","last_title":""})
        g["time_spent"] += ev.get("interaction_seconds") or 0
        ts = ev.get("created_at")
        if ts:
            if not g["first"] or ts < g["first"]: g["first"] = ts
            if not g["last"] or ts > g["last"]:
                g["last"] = ts
                g["ua"] = ev.get("user_agent") or g["ua"]
                g["ip"] = ev.get("remote_ip") or g["ip"]
                g["last_type"] = (ev.get("asset_type") or ev.get("asset_category") or "").title()
                url = ev.get("url") or ""
                m = re.search(r'/pages/([^/?#]+)', url)
                if m: g["last_title"] = f"Page: {m.group(1).replace('-',' ').title()}"
                else:
                    m = re.search(r'/modules/items/(\d+)', url)
                    g["last_title"] = f"Module Item {m.group(1)}" if m else (url[-80:] if url else "")
    return per_course

def courses_from_pageviews_filtered_to_account_tree(sess, user_id, target_account_id, since_iso):
    """
    Fallback: deduce cursos desde Page Views y conserva los que pertenezcan al árbol de la subcuenta 2070.
    """
    per_course = page_views_by_course(sess, user_id, since_iso=since_iso)
    course_ids = []
    for cid in sorted(per_course.keys()):
        cinfo = get_course_info(sess, cid)
        acc = cinfo.get("account_id")
        if is_in_account_tree(sess, acc, target_account_id):
            course_ids.append(cid)
    return course_ids, per_course

def last_submission_via_analytics(sess, course_id, user_id):
    r = sess.get(f"{API}/courses/{course_id}/analytics/users/{user_id}/assignments", timeout=30)
    if r.status_code != 200: return None
    last = None
    for a in (r.json() or []):
        sub = a.get("submission") or {}
        sa = sub.get("submitted_at")
        if sa and (not last or sa > last): last = sa
    return last

# ------------- impresión -------------
def print_table(rows, title):
    headers = ["Course ID","Curso","Estado","Rol","Tiempo","Primer acceso","Último acceso",
               "Última actividad","IP","Navegador/SO","Última entrega"]
    widths  = [10,36,10,18,8,19,19,30,15,18,19]
    line = "+" + "+".join("-"*w for w in widths) + "+"
    print(f"\n{title}\n")
    print(line); print("|" + "|".join(h.center(w) for h, w in zip(headers, widths)) + "|"); print(line)
    for r in rows:
        cells = [
            str(r["course_id"]),
            r["course_name"][:widths[1]-1],
            r["state"] or "-",
            r["role"] or "-",
            r["time"],
            r["first"] or "-",
            r["last"] or "-",
            (r["last_type"] + (" — " if r["last_title"] else "") + r["last_title"]).strip(" — ")[:widths[7]-1],
            (r["ip"] or "-")[:widths[8]-1],
            (r["browser"] + " / " + r["os"])[:widths[9]-1],
            r["last_sub"] or "-"
        ]
        print("|" + "|".join(f" {c:<{w-1}}" for c, w in zip(cells, widths)) + "|")
    print(line)

# ------------- main -------------
def main():
    sis_code = (sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SIS)
    since_iso = os.getenv("SINCE_ISO") or f"{dt.datetime.utcnow().year}-01-01T00:00:00Z"

    sess = session_or_die()

    # 1) user_id por SIS
    user_id, user_obj = get_user_id_by_sis(sess, sis_code)
    if not user_id:
        print(f"No encontré usuario con SIS '{sis_code}'. Revisa permiso 'Read SIS data'."); sys.exit(1)

    # 2) cursos de la subcuenta 2070 (incluye subcuentas) vía ruta ADMIN
    try:
        courses = list_courses_for_user_in_account(sess, ACCOUNT_ID, user_id)
        # FILTRO: quedarnos con cursos cuyo account_id esté en el árbol de 2070 (por seguridad)
        course_ids = []
        for c in courses:
            cid, acc = c.get("id"), c.get("account_id")
            if cid and is_in_account_tree(sess, acc, ACCOUNT_ID):
                course_ids.append(cid)
        print(f"✅ Cursos obtenidos vía /accounts/{ACCOUNT_ID}/courses: {len(course_ids)}")
        per_course = page_views_by_course(sess, user_id, since_iso=since_iso)  # métricas
    except requests.HTTPError:
        print(f"⚠️ /accounts/{ACCOUNT_ID}/courses devolvió 403 o vacío. Fallback por Page Views…")
        course_ids, per_course = courses_from_pageviews_filtered_to_account_tree(sess, user_id, ACCOUNT_ID, since_iso)
        if not course_ids:
            print("❌ No pude inferir cursos en esa subcuenta (ni subcuentas) desde Page Views en el rango dado.")
            sys.exit(1)
        print(f"✅ Cursos deducidos (subcuenta {ACCOUNT_ID} + subcuentas) desde Page Views: {len(course_ids)}")

    # 3) armar filas
    rows = []
    for cid in course_ids:
        cinfo = get_course_info(sess, cid)
        if not is_in_account_tree(sess, cinfo.get("account_id"), ACCOUNT_ID):
            continue  # seguridad doble
        enr  = get_user_enrollment_in_course_admin(sess, cid, user_id)
        info = per_course.get(cid, {})
        br, osn = parse_ua(info.get("ua",""))
        last_sub = iso_human(last_submission_via_analytics(sess, cid, user_id))
        rows.append({
            "course_id": cid,
            "course_name": cinfo["name"],
            "state": enr.get("state"),
            "role": enr.get("role"),
            "time": hms(info.get("time_spent", 0)),
            "first": iso_human(info.get("first")),
            "last":  iso_human(info.get("last")),
            "last_type": info.get("last_type",""),
            "last_title": info.get("last_title",""),
            "ip": info.get("ip",""),
            "browser": br,
            "os": osn,
            "last_sub": last_sub if last_sub != "-" else "-"
        })

    title = f"Alumno SIS: {sis_code} — {user_obj.get('name','(sin nombre)')}  |  Subcuenta: {ACCOUNT_ID} (incluye subcuentas)"
    print_table(rows, title)

if __name__ == "__main__":
    main()
