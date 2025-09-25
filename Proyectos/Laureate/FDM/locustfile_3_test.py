# locustfile.py — Genera token si falta, guarda en token.txt, recarga en 401 + stress duro
from locust import HttpUser, task, constant, events
import os, random, json, requests, pandas as pd
from pathlib import Path

# ================== CONFIG ==================
BASE = "https://apicert-manager.upc.edu.pe"
TOKEN_PATH = "/seguridad/v4.0/GenerarToken"
FDM_PATH = "/Academico/v4.0/AvanceFdm"  # con '/'

# Archivo donde guardamos/recargamos el token (igual que tu base)
TOKEN_FILE = Path(r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\token.txt")

# Si algún día la generación exigiera un Bearer previo, defínelo como env var:
#   PowerShell:  $env:PRE_AUTH_TOKEN="eyJ..."
PRE_AUTH_TOKEN = os.getenv("PRE_AUTH_TOKEN", "").strip()

# Excel (igual que tu base)
EXCEL_PATH  = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET       = 0
PIDM_COL    = "PIDM"
PROGRAM_COL = "COD_PROGRAMA"

# Stress
REQS_PER_CYCLE = 3
WAIT_TIME_SEC  = 0
WARMUP_TIMEOUT = 15  # s (para validar token una vez)
# ============================================

pairs = []
default_pair = ("EMA_1016_1V1", "1064861")

# ---------- Token helpers ---------- #
def _write_atomically(path: Path, text: str):
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    tmp.replace(path)

def _fetch_token() -> str:
    """Llama a /GenerarToken (con o sin Authorization) y devuelve accessToken."""
    url = f"{BASE}{TOKEN_PATH}"
    headers = {"Accept": "application/json"}
    if PRE_AUTH_TOKEN:
        headers["Authorization"] = f"Bearer {PRE_AUTH_TOKEN}"
    r = requests.get(url, headers=headers, timeout=20)
    r.raise_for_status()
    try:
        data = r.json()
    except json.JSONDecodeError:
        raise RuntimeError(f"Respuesta no JSON al generar token: {r.text[:300]}")
    tok = data.get("accessToken") or data.get("token") or data.get("access_token")
    if not tok:
        raise RuntimeError(f"No encontré 'accessToken' en la respuesta: {data}")
    return tok.strip()

def _ensure_token_on_disk() -> str:
    """
    Si token.txt no existe o está vacío → genera y guarda token.
    Devuelve el token actual (desde disco).
    """
    token = ""
    if TOKEN_FILE.exists():
        token = TOKEN_FILE.read_text(encoding="utf-8").strip()
    if not token:
        token = _fetch_token()
        _write_atomically(TOKEN_FILE, token)
        print(f"[Locust] Token generado y guardado en {TOKEN_FILE} (len={len(token)})")
    else:
        print(f"[Locust] Token listo (len={len(token)}) desde {TOKEN_FILE}")
    return token

def _regenerate_token_on_disk() -> str:
    token = _fetch_token()
    _write_atomically(TOKEN_FILE, token)
    print(f"[Locust] Token regenerado (len={len(token)})")
    return token

def _read_token() -> str:
    if not TOKEN_FILE.exists():
        return _ensure_token_on_disk()
    tok = TOKEN_FILE.read_text(encoding="utf-8").strip()
    return tok if tok else _ensure_token_on_disk()

# ---------- Excel helpers (como tu base) ---------- #
def load_pairs(excel_path: str, sheet, pidm_col: str, program_col: str):
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype={pidm_col: str, program_col: str})
    df = df[[pidm_col, program_col]].dropna()
    df[pidm_col]    = df[pidm_col].astype(str).str.extract(r'(\d+)', expand=False)
    df[program_col] = df[program_col].astype(str).str.strip()
    df = df[(df[pidm_col].str.len() > 0) & (df[program_col].str.len() > 0)].drop_duplicates([pidm_col, program_col])
    return list(df.apply(lambda r: (r[program_col], r[pidm_col]), axis=1))

# ---------- Warm-up (opcional pero útil) ---------- #
def _warmup_or_fail():
    """Valida el token una vez con un GET real para fallar rápido si es inválido."""
    tok = _read_token()
    headers = {"Authorization": f"Bearer {tok}", "Accept": "application/json"}
    cod, pidm = default_pair
    r = requests.get(f"{BASE}{FDM_PATH}",
                     params={"codigoPrograma": cod, "pidm": pidm},
                     headers=headers, timeout=WARMUP_TIMEOUT)
    if r.status_code == 401:
        # caducó antes de arrancar → regenerar y reintentar 1 vez
        new_tok = _regenerate_token_on_disk()
        headers["Authorization"] = f"Bearer {new_tok}"
        r = requests.get(f"{BASE}{FDM_PATH}",
                         params={"codigoPrograma": cod, "pidm": pidm},
                         headers=headers, timeout=WARMUP_TIMEOUT)
    r.raise_for_status()

# ---------- Hooks ---------- #
@events.test_start.add_listener
def on_test_start(environment, **kwargs):
    # 1) Asegura/genera token en disco y warm-up
    _ensure_token_on_disk()
    _warmup_or_fail()
    print("[Locust] Warm-up OK (token válido).")

    # 2) Carga Excel
    global pairs
    try:
        pairs = load_pairs(EXCEL_PATH, SHEET, PIDM_COL, PROGRAM_COL)
        if not pairs:
            pairs = [default_pair]
            print("[Locust] Excel sin filas válidas; usando par por defecto.")
        else:
            print(f"[Locust] Pares cargados: {len(pairs)}")
    except Exception as e:
        pairs = [default_pair]
        print(f"[Locust] No se pudo leer el Excel ({e}). Usaré par por defecto.")

# ---------- User ---------- #
class AvanceFdmUser(HttpUser):
    wait_time = constant(WAIT_TIME_SEC)

    def on_start(self):
        self.headers = {
            "Authorization": f"Bearer {_read_token()}",
            "Accept": "application/json"
        }

    @task
    def avance_fdm(self):
        for _ in range(REQS_PER_CYCLE):
            codigoPrograma, pidm = random.choice(pairs)
            params = {"codigoPrograma": codigoPrograma, "pidm": pidm}

            with self.client.get(FDM_PATH, params=params, headers=self.headers, catch_response=True) as resp:
                if resp.status_code == 401:
                    # caducó durante la corrida → regenerar y reintentar 1 vez
                    new_tok = _regenerate_token_on_disk()
                    self.headers["Authorization"] = f"Bearer {new_tok}"
                    retry = self.client.get(FDM_PATH, params=params, headers=self.headers)
                    if retry.status_code != 200:
                        resp.failure(f"401 reintento fallido: {retry.status_code} → PIDM={pidm} Programa={codigoPrograma}")
                    else:
                        resp.success()
                elif resp.status_code != 200:
                    resp.failure(f"{resp.status_code} → PIDM={pidm} Programa={codigoPrograma}")
                else:
                    resp.success()
