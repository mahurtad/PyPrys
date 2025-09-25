# locustfile.py — Lee Bearer desde token.txt + carga Excel + stress duro
from locust import HttpUser, task, constant, events
import random, pandas as pd
from pathlib import Path

# ================== CONFIG ==================
BASE = "https://apicert-manager.upc.edu.pe"
FDM_PATH = "/Academico/v4.0/AvanceFdm"  # con '/'

# Donde guardamos el token con generate_token.py
TOKEN_FILE = Path(r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\token.txt")

# Excel
EXCEL_PATH  = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET       = 0
PIDM_COL    = "PIDM"
PROGRAM_COL = "COD_PROGRAMA"

# Stress
REQS_PER_CYCLE = 3
WAIT_TIME_SEC  = 0
# ============================================

pairs = []
default_pair = ("EMA_1016_1V1", "1064861")

def get_token() -> str:
    if not TOKEN_FILE.exists():
        raise RuntimeError(f"No encontré el archivo de token: {TOKEN_FILE}")
    tok = TOKEN_FILE.read_text(encoding="utf-8").strip()
    if not tok:
        raise RuntimeError(f"El archivo {TOKEN_FILE} está vacío.")
    return tok

def load_pairs(excel_path: str, sheet, pidm_col: str, program_col: str):
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype={pidm_col: str, program_col: str})
    df = df[[pidm_col, program_col]].dropna()
    df[pidm_col]    = df[pidm_col].astype(str).str.extract(r'(\d+)', expand=False)
    df[program_col] = df[program_col].astype(str).str.strip()
    df = df[(df[pidm_col].str.len() > 0) & (df[program_col].str.len() > 0)].drop_duplicates([pidm_col, program_col])
    return list(df.apply(lambda r: (r[program_col], r[pidm_col]), axis=1))

@events.test_start.add_listener
def on_test_start(environment, **kwargs):
    # 1) Verifica token
    tok = get_token()
    print(f"[Locust] Token listo (len={len(tok)}). Fuente: {TOKEN_FILE}")

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

class AvanceFdmUser(HttpUser):
    wait_time = constant(WAIT_TIME_SEC)

    def on_start(self):
        self.headers = {
            "Authorization": f"Bearer {get_token()}",
            "Accept": "application/json"
        }

    @task
    def avance_fdm(self):
        for _ in range(REQS_PER_CYCLE):
            codigoPrograma, pidm = random.choice(pairs)
            params = {"codigoPrograma": codigoPrograma, "pidm": pidm}
            with self.client.get(FDM_PATH, params=params, headers=self.headers, catch_response=True) as resp:
                if resp.status_code == 401:
                    # Si regeneras el token durante la corrida, lo re-leemos y reintentamos una vez
                    self.headers["Authorization"] = f"Bearer {get_token()}"
                    retry = self.client.get(FDM_PATH, params=params, headers=self.headers)
                    if retry.status_code != 200:
                        resp.failure(f"401 reintento fallido: {retry.status_code} → PIDM={pidm} Programa={codigoPrograma}")
                    else:
                        resp.success()
                elif resp.status_code != 200:
                    resp.failure(f"{resp.status_code} → PIDM={pidm} Programa={codigoPrograma}")
                else:
                    resp.success()
