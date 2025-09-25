# locustfile.py — Stress duro (máximo RPS)
from locust import HttpUser, task
from locust import constant
from locust import events
import random
import pandas as pd

# ================== CONFIG ==================
BASE = "https://apicert-manager.upc.edu.pe"
FDM_PATH = "/Academico/v4.0/AvanceFdm"

# ► Token HARDCODEADO (reemplazar por el real)
API_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJDbGFpbVR5cGVzLlVzZXJuYW1lIjoidXNlci5Vc2VybmFtZSIsIkNsYWltVHlwZXMuQ29kaWdvVXN1YXJpbyI6InVzZXIuQ29kaWdvVXN1YXJpbyIsIkNsYWltVHlwZXMuRW1haWwiOiJ1c2VyLkVtYWlsIiwianRpIjoiMzBhMDkxN2YtMzMyYy00MGRkLWEzMDItMTYwYjU1NmQ2ODVlIiwiaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS93cy8yMDA4LzA2L2lkZW50aXR5L2NsYWltcy9yb2xlIjoiVXNlckFwcCIsImV4cCI6MTc1NzUyMjc1M30.CWlMnjQJOrstOada1YN7pon9gRL64vLueSTvN1fGn3Y"   # <-- pega aquí tu token válido

# ► Mismos parámetros del Excel (reutiliza tu script)
EXCEL_PATH  = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET       = 0
PIDM_COL    = "PIDM"
PROGRAM_COL = "COD_PROGRAMA"

# ► Cantidad de requests que hará cada usuario por ciclo (sube para más RPS)
REQS_PER_CYCLE = 3
# ===========================================

# Se cargan una sola vez para toda la prueba
pairs = []          # lista de tuplas (codigoPrograma, pidm)
default_pair = ("EMA_1016_1V1", "1064861")

def load_pairs(excel_path: str, sheet, pidm_col: str, program_col: str):
    """Igual que tu script: limpia y devuelve (COD_PROGRAMA, PIDM)."""
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype={pidm_col: str, program_col: str})
    df = df[[pidm_col, program_col]].dropna()

    df[pidm_col] = df[pidm_col].astype(str).str.extract(r'(\d+)', expand=False)  # solo dígitos
    df[program_col] = df[program_col].astype(str).str.strip()

    df = df[(df[pidm_col].str.len() > 0) & (df[program_col].str.len() > 0)]
    df = df.drop_duplicates([pidm_col, program_col])
    return list(df.apply(lambda r: (r[program_col], r[pidm_col]), axis=1))

@events.test_start.add_listener
def on_test_start(environment, **kwargs):
    """Lee el Excel una única vez al iniciar la corrida."""
    global pairs
    try:
        pairs = load_pairs(EXCEL_PATH, SHEET, PIDM_COL, PROGRAM_COL)
        if not pairs:
            pairs = [default_pair]
            print("[Locust] Excel leído pero sin filas válidas; usando par por defecto.")
        else:
            print(f"[Locust] Pares cargados: {len(pairs)}")
    except Exception as e:
        pairs = [default_pair]
        print(f"[Locust] No se pudo leer el Excel ({e}). Usaré par por defecto.")

class AvanceFdmUser(HttpUser):
    # Máximo throughput: sin “think time”
    wait_time = constant(0)

    def on_start(self):
        self.headers = {
            "Authorization": f"Bearer {API_TOKEN}",
            "Accept": "application/json"
        }

    @task
    def avance_fdm(self):
        # Cada ciclo hará varias requests para multiplicar el RPS
        for _ in range(REQS_PER_CYCLE):
            codigoPrograma, pidm = random.choice(pairs)
            params = {"codigoPrograma": codigoPrograma, "pidm": pidm}

            with self.client.get(FDM_PATH, params=params, headers=self.headers, catch_response=True) as resp:
                if resp.status_code != 200:
                    resp.failure(f"{resp.status_code} → PIDM={pidm} Programa={codigoPrograma}")
                else:
                    resp.success()
