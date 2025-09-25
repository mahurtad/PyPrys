#!/usr/bin/env python
# -*- coding: utf-8 -*-
r"""
One-shot: lee el Excel, genera token y llama AvanceFdm para cada fila,
exportando a Excel con hoja de requests (duración por request) y summary.

Requisitos:
  pip install httpx pandas openpyxl

Ejecución:
  python fdm_service.py
"""

# ============================ CONFIG ============================

EXCEL_PATH   = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET        = 0
PIDM_COL     = "PIDM"
PROGRAM_COL  = "COD_PROGRAMA"

# Credenciales para /seguridad/v4.0/GenerarToken (Basic Auth)
BASIC_USER   = r"DOMINIO\usuario"  # deja "" para que te lo pida
BASIC_PASS   = ""                  # deja "" para que te lo pida (oculto)

# Concurrencia y timeouts
CONCURRENCY  = 20
TIMEOUT_SEC  = 15.0

# Endpoints base
BASE         = "https://apicert-manager.upc.edu.pe"
TOKEN_PATH   = "/seguridad/v4.0/GenerarToken"
FDM_PATH     = "/Academico/v4.0/AvanceFdm"

# ===============================================================

import asyncio
import getpass
import os
import statistics
import time
from datetime import datetime
from typing import List, Tuple, Optional, Dict, Any

import httpx
import pandas as pd


# ----------------------------- Utilidades ----------------------------- #

def ts_now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def build_output_path_xlsx(excel_path: str) -> str:
    p = os.path.abspath(excel_path)
    base_dir = os.path.dirname(p)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(base_dir, f"fdm_results_{ts}.xlsx")

def percentile(sorted_vals: List[float], p: float) -> float:
    if not sorted_vals:
        return 0.0
    k = (len(sorted_vals) - 1) * p
    f = int(k)
    c = min(f + 1, len(sorted_vals) - 1)
    if f == c:
        return sorted_vals[f]
    return sorted_vals[f] * (c - k) + sorted_vals[c] * (k - f)

def load_pairs(excel_path: str, sheet, pidm_col: str, program_col: str) -> List[Tuple[str, str]]:
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype={pidm_col: str, program_col: str})
    df = df[[pidm_col, program_col]].dropna()
    # Limpieza básica
    df[pidm_col] = df[pidm_col].astype(str).str.extract(r'(\d+)', expand=False)  # solo dígitos
    df[program_col] = df[program_col].astype(str).str.strip()
    df = df[(df[pidm_col].str.len() > 0) & (df[program_col].str.len() > 0)]
    df = df.drop_duplicates([pidm_col, program_col])
    # Tuplas (codigoPrograma, pidm)
    return list(df.apply(lambda r: (r[program_col], r[pidm_col]), axis=1))


# ----------------------------- Token Manager ----------------------------- #

class TokenManager:
    """Obtiene y refresca el accessToken (refresca una vez si hay 401)."""
    def __init__(self, user: str, pwd: str, timeout: float):
        self.user = user
        self.pwd = pwd
        self.timeout = timeout
        self._token: Optional[str] = None
        self._lock = asyncio.Lock()

    async def _fetch_new(self) -> str:
        url = BASE + TOKEN_PATH
        async with httpx.AsyncClient(timeout=self.timeout, http2=True) as client:
            r = await client.get(url, auth=(self.user, self.pwd))
            r.raise_for_status()
            data = r.json()
            tok = data.get("accessToken")
            if not tok:
                raise RuntimeError("Token response missing 'accessToken'")
            return tok

    async def get(self) -> str:
        if self._token:
            return self._token
        async with self._lock:
            if not self._token:
                self._token = await self._fetch_new()
            return self._token

    async def refresh(self) -> str:
        async with self._lock:
            self._token = await self._fetch_new()
            return self._token


# ----------------------------- HTTP Lógica ----------------------------- #

async def fetch_one(client: httpx.AsyncClient, token_mgr: TokenManager, program: str, pidm: str) -> Dict[str, Any]:
    params = {"codigoPrograma": program, "pidm": pidm}
    started_at = ts_now()
    t0 = time.perf_counter()
    token = await token_mgr.get()
    row: Dict[str, Any] = {
        "codigoPrograma": program,
        "pidm": pidm,
        "started_at": started_at,
    }
    try:
        resp = await client.get(FDM_PATH, params=params, headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        })
        status = resp.status_code
        if status == 401:
            token = await token_mgr.refresh()
            resp = await client.get(FDM_PATH, params=params, headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json"
            })
            status = resp.status_code

        elapsed_ms = (time.perf_counter() - t0) * 1000.0
        finished_at = ts_now()

        # Longitud de body (referencial)
        try:
            _ = resp.json()
            body_len = len(resp.text or "")
        except Exception:
            body_len = len(resp.text or "")

        row.update({
            "status": status,
            "ok": status in (200, 204),
            "duration_ms": round(elapsed_ms, 2),
            "finished_at": finished_at,
            "len_body": body_len
        })
        return row

    except Exception as e:
        elapsed_ms = (time.perf_counter() - t0) * 1000.0
        finished_at = ts_now()
        row.update({
            "status": -1,
            "ok": False,
            "duration_ms": round(elapsed_ms, 2),
            "finished_at": finished_at,
            "error": str(e)[:200]
        })
        return row


async def run_batch(pairs: List[Tuple[str, str]], token_mgr: TokenManager, concurrency: int, timeout: float):
    limits = httpx.Limits(max_connections=concurrency, max_keepalive_connections=concurrency)
    async with httpx.AsyncClient(base_url=BASE, timeout=timeout, limits=limits, http2=True) as client:
        sem = asyncio.Semaphore(concurrency)
        results: List[Dict[str, Any]] = []

        async def bound(pgm: str, pidm: str):
            async with sem:
                res = await fetch_one(client, token_mgr, pgm, pidm)
                results.append(res)

        t_start = time.perf_counter()
        tasks = [asyncio.create_task(bound(p, m)) for (p, m) in pairs]
        await asyncio.gather(*tasks)
        total_seconds = time.perf_counter() - t_start
        return results, total_seconds


# ----------------------------- Reporte Excel ----------------------------- #

def write_excel_report(rows: List[Dict[str, Any]], total_seconds: float, out_xlsx: str) -> None:
    df = pd.DataFrame(rows)

    # Métricas
    total = len(rows)
    ok = int(df["ok"].sum()) if "ok" in df else 0
    err = total - ok
    status_counts = df["status"].value_counts(dropna=False).sort_index() if "status" in df else pd.Series(dtype=int)
    lats = sorted(df.get("duration_ms", pd.Series(dtype=float)).dropna().tolist())

    avg = statistics.mean(lats) if lats else 0.0
    p50 = percentile(lats, 0.50)
    p90 = percentile(lats, 0.90)
    p95 = percentile(lats, 0.95)
    p99 = percentile(lats, 0.99)
    mx  = max(lats) if lats else 0.0
    rps = (total / total_seconds) if total_seconds > 0 else 0.0

    # Summary como DataFrame vertical
    summary_rows = [
        {"metric": "total_requests", "value": total},
        {"metric": "ok", "value": ok},
        {"metric": "errors", "value": err},
        {"metric": "elapsed_seconds", "value": round(total_seconds, 3)},
        {"metric": "throughput_rps", "value": round(rps, 3)},
        {"metric": "avg_ms", "value": round(avg, 2)},
        {"metric": "p50_ms", "value": round(p50, 2)},
        {"metric": "p90_ms", "value": round(p90, 2)},
        {"metric": "p95_ms", "value": round(p95, 2)},
        {"metric": "p99_ms", "value": round(p99, 2)},
        {"metric": "max_ms", "value": round(mx, 2)},
    ]
    for s, cnt in status_counts.items():
        summary_rows.append({"metric": f"status_{s}", "value": int(cnt)})

    df_summary = pd.DataFrame(summary_rows)

    # Escribir Excel
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        cols = ["codigoPrograma", "pidm", "status", "ok",
                "duration_ms", "started_at", "finished_at", "len_body", "error"]
        cols = [c for c in cols if c in df.columns] + [c for c in df.columns if c not in cols]
        df[cols].to_excel(writer, index=False, sheet_name="requests")
        df_summary.to_excel(writer, index=False, sheet_name="summary")


# ----------------------------- main ----------------------------- #

def main():
    # Pide credenciales si no están en CONFIG
    user = BASIC_USER.strip() if BASIC_USER else input(r"Basic user (DOMINIO\usuario): ").strip()
    pwd  = BASIC_PASS if BASIC_PASS else getpass.getpass("Basic password: ")

    if not os.path.isfile(EXCEL_PATH):
        raise FileNotFoundError(f"No existe el Excel: {EXCEL_PATH}")

    pairs = load_pairs(EXCEL_PATH, SHEET, PIDM_COL, PROGRAM_COL)
    if not pairs:
        print("No se encontraron pares (COD_PROGRAMA, PIDM) en el Excel.")
        return
    print(f"Pares a procesar: {len(pairs)}  |  Concurrencia: {CONCURRENCY}")

    token_mgr = TokenManager(user, pwd, TIMEOUT_SEC)
    rows, total_seconds = asyncio.run(run_batch(pairs, token_mgr, CONCURRENCY, TIMEOUT_SEC))

    out_xlsx = build_output_path_xlsx(EXCEL_PATH)
    write_excel_report(rows, total_seconds, out_xlsx)

    print(f"Reporte Excel generado: {out_xlsx}")
    print(f"Total filas: {len(rows)}  |  Duración total: {round(total_seconds, 3)} s")


if __name__ == "__main__":
    main()
