#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
from urllib.parse import urljoin
import requests
from bs4 import BeautifulSoup

BASE = "http://192.168.1.1/"
LOGIN_PATH = "cgi-bin/logIn_mhs.cgi"
NETWORK_MAP_PATH = "cgi-bin/indexmain.cgi"

ROUTER_PASSWORD = "#EDCC27@DC48AA*"   # <-- no la compartas
TIMEOUT = 8

MAC_RE = re.compile(r"([0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2}")
IP_RE  = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")

# fabricante (OUI)
try:
    from mac_vendor_lookup import MacLookup
    _mac = MacLookup(); _mac.load_vendors()
    def vendor(mac):
        try: return _mac.lookup(mac)
        except Exception: return ""
except Exception:
    def vendor(mac): return ""

def _get(s, path):
    r = s.get(urljoin(BASE, path), timeout=TIMEOUT, allow_redirects=True)
    r.raise_for_status(); return r

def _post(s, path, data):
    r = s.post(urljoin(BASE, path), data=data, timeout=TIMEOUT, allow_redirects=True)
    r.raise_for_status(); return r

def login(session: requests.Session) -> bool:
    """GET login -> parsea form -> POST con hidden + contraseña."""
    session.headers.update({
        "User-Agent": "Mozilla/5.0",
        "Referer": urljoin(BASE, LOGIN_PATH),
        "Origin": BASE.rstrip("/"),
    })

    # 1) GET para obtener cookies/tokens
    r = _get(session, LOGIN_PATH)
    soup = BeautifulSoup(r.text, "html.parser")

    form = None
    for f in soup.find_all("form"):
        if f.find("input", {"type": "password"}):
            form = f; break
    if not form:
        # si no encontramos form, probamos un POST simple
        payload = {"syspasswd": ROUTER_PASSWORD, "leaveBlur": "0", "Submit": "Comprobar"}
        _post(session, LOGIN_PATH, payload)
    else:
        # 2) construir payload con todos los hidden y los campos de password
        payload = {}
        for inp in form.find_all("input"):
            name = inp.get("name")
            if not name: continue
            t = (inp.get("type") or "").lower()
            if t in ("submit", "button"):  # no hace falta
                continue
            payload[name] = inp.get("value", "")
        # poner contraseña en ambos por si el JS duplica campos
        payload["syspasswd"] = ROUTER_PASSWORD
        payload["fake_syspasswd"] = ROUTER_PASSWORD
        # algunos firmwares usan un segundo campo
        payload["syspasswd_1"] = ROUTER_PASSWORD
        payload.setdefault("leaveBlur", "0")
        payload.setdefault("Submit", "Comprobar")

        action = form.get("action") or LOGIN_PATH
        _post(session, action, payload)

    # 3) comprobar acceso
    chk = _get(session, "cgi-bin/deviceinfo.cgi").text
    return ("logIn" not in chk) and ("window.parent.location.href" not in chk)

def parse_network_map(html: str):
    soup = BeautifulSoup(html, "html.parser")
    devices = []
    # localizar tabla que contenga 'Device Name' y 'MAC Address'
    for tbl in soup.find_all("table"):
        head = tbl.get_text(" ", strip=True).lower()
        if "device name" in head and "mac address" in head:
            for tr in tbl.find_all("tr"):
                cells = [td.get_text(" ", strip=True) for td in tr.find_all(["td","th"])]
                text = " | ".join(cells)
                m_mac = MAC_RE.search(text); m_ip = IP_RE.search(text)
                if not (m_mac and m_ip): continue
                name = ""
                # normalmente la 2ª columna es Device Name
                if len(cells) >= 2: name = cells[1]
                name = "" if name.lower() in ("unknown","n/a") else name
                mac = m_mac.group(0).lower().replace("-", ":")
                ip  = m_ip.group(0)
                devices.append((ip, mac, name))
    # dedup
    out, seen = [], set()
    for ip, mac, name in devices:
        if (ip, mac) in seen: continue
        seen.add((ip, mac)); out.append((ip, mac, name))
    return out

def main():
    if not ROUTER_PASSWORD or ROUTER_PASSWORD.startswith("<PON_"):
        print("⚠️  Edita la variable ROUTER_PASSWORD en el script.")
        return

    with requests.Session() as s:
        if not login(s):
            print("❌ No se pudo iniciar sesión (contraseña incorrecta, o faltan cookies/tokens).")
            print("Abre F12→Network y confirma que el POST es a /cgi-bin/logIn_mhs.cgi; si ves algún token no vacío (p. ej. sessionKey), dímelo para incluirlo.")
            return

        # Network Map
        html = _get(s, NETWORK_MAP_PATH).text
        devices = parse_network_map(html)
        if not devices:
            print("No se pudo extraer la tabla del Network Map.")
            return

        print(f"{'IP':<15} {'MAC':<17} {'Fabricante':<28} Nombre")
        print("-"*90)
        for ip, mac, name in devices:
            vend = vendor(mac)
            print(f"{ip:<15} {mac:<17} {vend[:27]:<28} {name}")
        print(f"\nTotal: {len(devices)} dispositivo(s).")

if __name__ == "__main__":
    main()
