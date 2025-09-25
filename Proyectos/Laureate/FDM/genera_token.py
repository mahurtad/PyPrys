#!/usr/bin/env python3
import argparse
import os
import sys
import json
from pathlib import Path
import requests

BASE = "https://apicert-manager.upc.edu.pe"
TOKEN_PATH = "/seguridad/v4.0/GenerarToken"

# Ruta destino solicitada
DEFAULT_TOKEN_FILE = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\token.txt"

def mask(s: str, left=6, right=4):
    if not s:
        return "<empty>"
    if len(s) <= left + right:
        return "*" * len(s)
    return f"{s[:left]}...{s[-right:]}"

def fetch_token(pre_auth: str | None, timeout: int = 15, insecure: bool = False) -> str:
    url = f"{BASE}{TOKEN_PATH}"
    headers = {"Accept": "application/json"}
    if pre_auth:
        headers["Authorization"] = f"Bearer {pre_auth}"

    resp = requests.get(url, headers=headers, timeout=timeout, verify=not insecure)
    try:
        resp.raise_for_status()
    except requests.HTTPError as e:
        # imprime cuerpo para ayudar a depurar
        raise RuntimeError(f"HTTP {resp.status_code} en GenerarToken: {resp.text[:400]}") from e

    try:
        data = resp.json()
    except json.JSONDecodeError:
        raise RuntimeError(f"Respuesta no JSON de GenerarToken: {resp.text[:400]}")
    token = data.get("accessToken") or data.get("token") or data.get("access_token")
    if not token:
        raise RuntimeError(f"No encontré 'accessToken' en la respuesta: {data}")
    return token.strip()

def write_token_file(path: Path, token: str, no_newline: bool = True) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if no_newline:
        # Evita salto de línea final (útil si otros scripts leen tal cual)
        path.write_text(token, encoding="utf-8")
    else:
        with path.open("w", encoding="utf-8") as f:
            f.write(token + "\n")

def main():
    ap = argparse.ArgumentParser(
        description="Genera un token desde /seguridad/v4.0/GenerarToken y lo guarda en token.txt"
    )
    ap.add_argument("--out", default=DEFAULT_TOKEN_FILE,
                    help="Ruta del archivo token.txt de salida "
                         f"(default: {DEFAULT_TOKEN_FILE})")
    ap.add_argument("--pre-auth",
                    default=os.getenv("PRE_AUTH_TOKEN", "").strip(),
                    help="Bearer previo (si el endpoint lo requiere). "
                         "Si no lo pasas, se intenta sin Authorization. "
                         "También puede leerse de la variable PRE_AUTH_TOKEN.")
    ap.add_argument("--timeout", type=int, default=15, help="Timeout de la llamada (s)")
    ap.add_argument("--insecure", action="store_true",
                    help="No verificar certificado TLS (solo para pruebas locales)")
    args = ap.parse_args()

    out_path = Path(args.out)

    try:
        token = fetch_token(pre_auth=args.pre_auth or None,
                            timeout=args.timeout,
                            insecure=args.insecure)
    except Exception as e:
        print(f"[ERROR] No se pudo obtener el token: {e}", file=sys.stderr)
        sys.exit(2)

    try:
        write_token_file(out_path, token, no_newline=True)
    except Exception as e:
        print(f"[ERROR] No se pudo escribir el archivo {out_path}: {e}", file=sys.stderr)
        sys.exit(3)

    print(f"[OK] Token generado y guardado en: {out_path}")
    print(f"[OK] Longitud: {len(token)}")
    print(f"[OK] Preview: {mask(token)}")

if __name__ == "__main__":
    main()
