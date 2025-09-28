#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import ipaddress, platform, socket, subprocess, re, sys
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from mac_vendor_lookup import MacLookup
    mac_lookup = MacLookup()
    # si no quieres que intente actualizar online:
    mac_lookup.load_vendors()  # usa la DB incluida
except Exception:
    mac_lookup = None

IP_RE   = re.compile(r"\b(?:\d{1,3}\.){3}\d{1,3}\b")
MAC_RE  = re.compile(r"([0-9a-fA-F]{2}[:-]){5}[0-9a-fA-F]{2}")

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

def ping(ip):
    sysname = platform.system().lower()
    if sysname == "windows":
        cmd = ["ping", "-n", "1", "-w", "80", str(ip)]
    else:
        cmd = ["ping", "-c", "1", "-W", "1", str(ip)]
    try:
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        pass

def populate_arp(network_cidr):
    ips = [str(ip) for ip in ipaddress.ip_network(network_cidr, strict=False).hosts()]
    with ThreadPoolExecutor(max_workers=128) as ex:
        for _ in as_completed([ex.submit(ping, ip) for ip in ips]):
            pass

def parse_arp():
    # Windows / macOS / Linux: `arp -a` suele existir
    try:
        out = subprocess.check_output(["arp", "-a"], text=True, stderr=subprocess.DEVNULL)
    except Exception as e:
        print("Error al ejecutar `arp -a`:", e)
        return []
    devices = []
    for line in out.splitlines():
        ip_m = IP_RE.search(line)
        mac_m = MAC_RE.search(line)
        if not ip_m or not mac_m:
            continue
        ip  = ip_m.group(0)
        mac = mac_m.group(0).replace("-", ":").lower()
        devices.append((ip, mac))
    return devices

def is_special_address(ip):
    # filtra multicast/broadcast y 224.x/239.x/255.255.255.255 y .255 local
    try:
        ipaddr = ipaddress.ip_address(ip)
        return ipaddr.is_multicast or ipaddr.is_loopback or ipaddr.is_reserved or ipaddr.is_unspecified or ip.endswith(".255")
    except Exception:
        return True

def resolve_name(ip):
    # 1) reverse DNS
    try:
        host = socket.gethostbyaddr(ip)[0]
        if host and host != ip:
            return host
    except Exception:
        pass
    # 2) Windows: ping -a (a veces revela NetBIOS/mDNS)
    if platform.system().lower() == "windows":
        try:
            p = subprocess.run(["ping", "-a", "-n", "1", ip], capture_output=True, text=True, timeout=2)
            m = re.search(r"Haciendo ping a (.+?) \[", p.stdout) or re.search(r"Ping request could not find host (.+?)\.", p.stdout)
            if m:
                candidate = m.group(1).strip()
                if candidate and candidate != ip:
                    return candidate
        except Exception:
            pass
        # 3) nbtstat -A (NetBIOS)
        try:
            p = subprocess.run(["nbtstat", "-A", ip], capture_output=True, text=True, timeout=2)
            for line in p.stdout.splitlines():
                # nombre <00> UNIQUE ... o <20> UNIQUE
                if "<00>" in line or "<20>" in line:
                    name = line.split()[0]
                    if name and name != "Nombre":
                        return name
        except Exception:
            pass
    return ""

def vendor_from_mac(mac):
    if mac_lookup is None:
        return ""
    try:
        return mac_lookup.lookup(mac)
    except Exception:
        # si no está en la DB, intentamos con solo el OUI
        return ""

def main():
    # red por defecto: /24 de tu IP local
    if len(sys.argv) > 1:
        net = sys.argv[1]
    else:
        local_ip = get_local_ip()
        net = local_ip.rpartition(".")[0] + ".0/24"

    print(f"Escaneando (ping) la red {net} para poblar tabla ARP...")
    populate_arp(net)
    print("Parseando tabla ARP...\n")

    pairs = parse_arp()
    # enriquecer con fabricante y nombre
    enriched = []
    for ip, mac in pairs:
        if is_special_address(ip) or mac == "ff:ff:ff:ff:ff:ff":
            continue
        name = resolve_name(ip)
        vendor = vendor_from_mac(mac)
        enriched.append((ip, mac, vendor, name))

    # ordenar por IP
    def ip_key(t):
        try:
            return tuple(int(x) for x in t[0].split("."))
        except Exception:
            return (999, 999, 999, 999)

    enriched.sort(key=ip_key)

    # salida consola
    print(f"{'IP':<15} {'MAC':<17} {'Fabricante':<28} Nombre de dispositivo")
    print("-" * 80)
    for ip, mac, vendor, name in enriched:
        print(f"{ip:<15} {mac:<17} {vendor[:27]:<28} {name}")

    print(f"\nTotal: {len(enriched)} dispositivo(s).")

if __name__ == "__main__":
    main()
