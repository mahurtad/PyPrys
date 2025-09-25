# --- helpers de paginado y UA ---
import requests, datetime as dt, re

def fetch_all(session, url, params=None):
    params = params or {}
    params.setdefault('per_page', 100)
    out = []
    while url:
        r = session.get(url, params=params)
        r.raise_for_status()
        out.extend(r.json() if isinstance(r.json(), list) else [r.json()])
        # parse Link header
        link = r.headers.get('Link', '')
        m = re.search(r'<([^>]+)>;\s*rel="next"', link)
        url = m.group(1) if m else None
        params = None  # después del primer GET, la URL ya lleva querystring
    return out

def parse_ua(ua):
    # súper simple; si querés algo mejor: pip install httpagentparser
    b = "Unknown"
    if "Chrome" in ua and "Chromium" not in ua and "Edg" not in ua: b = "Chrome"
    elif "Firefox" in ua: b = "Firefox"
    elif "Edg" in ua: b = "Edge"
    elif "Safari" in ua and "Chrome" not in ua: b = "Safari"
    os = "Unknown"
    if "Windows" in ua: os = "Windows"
    elif "Mac OS X" in ua or "Macintosh" in ua: os = "macOS"
    elif "Android" in ua: os = "Android"
    elif "iPhone" in ua or "iPad" in ua: os = "iOS/iPadOS"
    elif "Linux" in ua: os = "Linux"
    return b, os

def iso(ts): 
    return ts and dt.datetime.fromisoformat(ts.replace('Z','+00:00')).isoformat()

# --- dataset del grid para un curso ---
def build_course_activity_dataset(base_api, session, course_id, since_iso=None, until_iso=None):
    # 1) roster (solo estudiantes)
    students = fetch_all(session, f"{base_api}/courses/{course_id}/users",
                         params={'enrollment_type[]':'student'})
    enrollments = fetch_all(session, f"{base_api}/courses/{course_id}/enrollments")
    enr_by_user = {}
    for e in enrollments:
        if e.get('type') == 'StudentEnrollment':
            enr_by_user.setdefault(e['user_id'], []).append(e)

    # 2) por estudiante: page views + métricas
    rows = []
    for u in students:
        uid = u['id']
        pv_url = f"{base_api}/users/{uid}/page_views"
        params = {}
        if since_iso: params['start_time'] = since_iso   # e.g. '2025-01-01T00:00:00Z'
        if until_iso: params['end_time']   = until_iso
        try:
            pvs = fetch_all(session, pv_url, params=params)
        except requests.HTTPError as e:
            # sin permisos en algún usuario → seguimos
            pvs = []

        time_spent = 0
        first_access = None
        last_access  = None
        last_ua = ""
        last_ip = ""
        last_activity_type = ""
        last_activity_title = ""

        for ev in pvs:
            time_spent += ev.get('interaction_seconds') or 0
            ts = ev.get('created_at')
            if ts:
                if not first_access or ts < first_access: first_access = ts
                if not last_access  or ts > last_access:
                    last_access = ts
                    last_ip = ev.get('remote_ip') or last_ip
                    last_ua = ev.get('user_agent') or last_ua
                    last_activity_type = (ev.get('asset_type') or ev.get('asset_category') or "").title()
                    # título/actividad: mejor esfuerzo desde la URL
                    url = (ev.get('url') or "")
                    # ejemplo: /courses/123/pages/lesson-3 → "Page: lesson-3"
                    m = re.search(r'/pages/([^/?#]+)', url)
                    if m: last_activity_title = f"Page: {m.group(1).replace('-', ' ').title()}"
                    else:
                        m = re.search(r'/modules/items/(\d+)', url)
                        last_activity_title = f"Module Item {m.group(1)}" if m else url[-80:]

        browser, osname = parse_ua(last_ua)

        # 3) última entrega (sin permisos de submissions → usamos Analytics como aproximación)
        last_submission = None
        try:
            anns = session.get(f"{base_api}/courses/{course_id}/analytics/users/{uid}/assignments")
            if anns.status_code == 200:
                data = anns.json() or []
                # algunos tenants incluyen submission.submitted_at aquí; si no, queda None
                for a in data:
                    sub = a.get('submission') or {}
                    sa = sub.get('submitted_at')
                    if sa and (not last_submission or sa > last_submission):
                        last_submission = sa
        except Exception:
            pass

        # enrollment status/role (primer match)
        enr = (enr_by_user.get(uid) or [{}])[0]
        rows.append({
            "first_name": u.get('name', '').split(' ')[0],
            "last_name":  ' '.join((u.get('name','') or '').split(' ')[1:]),
            "email": u.get('login_id') or u.get('sis_user_id') or "",
            "enrollment_status": enr.get('enrollment_state') or enr.get('workflow_state'),
            "course_role": enr.get('type'),
            "activity_type": last_activity_type,
            "activity": last_activity_title,
            "time_spent_seconds": int(time_spent),
            "first_access": iso(first_access),
            "last_access":  iso(last_access),
            "last_submission": iso(last_submission),
            "browser": browser,
            "operating_system": osname,
            "ip_address": last_ip
        })
    return rows
