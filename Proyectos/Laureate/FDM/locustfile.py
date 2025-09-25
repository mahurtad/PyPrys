# locustfile.py
from locust import HttpUser, task, between
import os, random
import pandas as pd

# ================== CONFIG ==================
BASE = "https://apicert-manager.upc.edu.pe"
FDM_PATH = "/Academico/v4.0/AvanceFdm"

# ► Token HARDCODEADO (reemplazar por el real)
API_TOKEN = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJDbGFpbVR5cGVzLlVzZXJuYW1lIjoidXNlci5Vc2VybmFtZSIsIkNsYWltVHlwZXMuQ29kaWdvVXN1YXJpbyI6InVzZXIuQ29kaWdvVXN1YXJpbyIsIkNsYWltVHlwZXMuRW1haWwiOiJ1c2VyLkVtYWlsIiwianRpIjoiNTExOTJhNTktYTQ4NS00OTkxLTlkODMtM2JlMGVmMmJmMjEyIiwiaHR0cDovL3NjaGVtYXMubWljcm9zb2Z0LmNvbS93cy8yMDA4LzA2L2lkZW50aXR5L2NsYWltcy9yb2xlIjoiVXNlckFwcCIsImV4cCI6MTc1NzQzNTk2MX0.2cuuBkf7SNcLWw1yLBnqnIGoYUcUjbPpDum3ARdCb9o"  # <-- pega tu token válido aquí

# ► Mismos parámetros de lectura que tu script original (Excel)
EXCEL_PATH  = r"G:\My Drive\Data Analysis\Proyectos\Laureate\FDM\szvfdmv_v2.xlsx"
SHEET       = 0
PIDM_COL    = "PIDM"
PROGRAM_COL = "COD_PROGRAMA"
# ============================================

def load_pairs(excel_path: str, sheet, pidm_col: str, program_col: str):
    """
    Carga (COD_PROGRAMA, PIDM) desde Excel y devuelve lista de tuplas.
    Misma limpieza básica que tu script: solo dígitos en PIDM, strip en programa,
    filtra vacíos y quita duplicados.
    """
    df = pd.read_excel(excel_path, sheet_name=sheet, dtype={pidm_col: str, program_col: str})
    df = df[[pidm_col, program_col]].dropna()

    df[pidm_col] = df[pidm_col].astype(str).str.extract(r'(\d+)', expand=False)  # solo dígitos
    df[program_col] = df[program_col].astype(str).str.strip()

    df = df[(df[pidm_col].str.len() > 0) & (df[program_col].str.len() > 0)]
    df = df.drop_duplicates([pidm_col, program_col])

    return list(df.apply(lambda r: (r[program_col], r[pidm_col]), axis=1))

class AvanceFdmUser(HttpUser):
    wait_time = between(1, 2)

    def on_start(self):
        # Carga una sola vez al iniciar cada usuario virtual
        self.headers = {
            "Authorization": f"Bearer {API_TOKEN}",
            "Accept": "application/json"
        }
        try:
            self.pairs = load_pairs(EXCEL_PATH, SHEET, PIDM_COL, PROGRAM_COL)
        except Exception as e:
            # Fallback simple si no encuentra el Excel
            self.pairs = [("EMA_1016_1V1", "1064861")]
            print(f"[Locust] Advertencia: no pude leer el Excel ({e}). Usaré un par por defecto.")

    @task
    def avance_fdm(self):
        codigoPrograma, pidm = random.choice(self.pairs)
        params = {"codigoPrograma": codigoPrograma, "pidm": pidm}

        # Usamos catch_response para marcar éxito/fracaso en el UI
        with self.client.get(FDM_PATH, params=params, headers=self.headers, catch_response=True) as resp:
            if resp.status_code != 200:
                resp.failure(f"{resp.status_code} para PIDM={pidm} Programa={codigoPrograma}")
            else:
                resp.success()
