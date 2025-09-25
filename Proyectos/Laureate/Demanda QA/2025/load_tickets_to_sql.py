# Reqs: pip install pandas sqlalchemy pyodbc openpyxl xlrd==1.2.0 pyxlsb

from pathlib import Path
import shutil
import pandas as pd
from sqlalchemy import create_engine, text

# ========= CONFIG =========
SQL_SERVER   = r"localhost\SQLEXPRESS"
SQL_DATABASE = "qa_liu_2025"
SQL_USERNAME = "mahurtad"
SQL_PASSWORD = "NICAMINICA"

EXCEL_PATH   = r"C:\Users\manue\OneDrive - EduCorpPERU\Proyectos QA - 2025\Gestión Demanda Certificaciones LIUv1.xlsx"
EXCEL_SHEET  = "Tickets Diarios"
LOCAL_COPY   = Path(r"C:\temp\liuv1.xlsx")   # copia local para evitar placeholders de OneDrive

# ========= CONEXIÓN =========
def make_engine(driver_name: str):
    odbc = (
        f"Driver={{{driver_name}}};"
        f"Server={SQL_SERVER};Database={SQL_DATABASE};UID={SQL_USERNAME};PWD={SQL_PASSWORD};"
        "TrustServerCertificate=yes;"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={odbc}", fast_executemany=True)

engine = None
for drv in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
    try:
        engine = make_engine(drv)
        with engine.connect() as _:
            pass
        print(f"✅ Conectado con: {drv}")
        break
    except Exception as e:
        print(f"⚠️ No se pudo con {drv}: {e}")
assert engine is not None, "No fue posible conectarse con ODBC 18 ni 17."

# ========= MAPEO Excel -> STAGING =========
COLUMN_MAP = {
    "INSTITUCION": "institucion",
    "RITM/INC": "RITM/INC",
    "SCTASK": "sctask",
    "DESCRIPCIÓN": "descripcion",
    "APLICATIVO/BD": "aplicativo_bd",
    "IMPACTO": "impacto",
    "RIESGO": "riesgo",
    "TRABAJO": "trabajo",
    "TIPO CAMBIO": "tipo_cambio",
    "SOLICITANTE": "solicitante",
    "AREA": "area",
    "DESARROLLADOR": "desarrollador",
    "AREA DESARROLLO": "area_desarrollo",
    "ANALISTA QA": "analista_qa",
    "FECHA RECEPCION COORDINADOR QA": "fecha_recepcion_coordinador_qa",
    "HORA RECEPCIÓN": "hora_recepcion",
    "FECHA ASIGNACIÓN": "fecha_asignacion",
    "HORA ASIGNACIÓN": "hora_asignacion",
    "CANT. FALLOS REPORTADOS": "cant_fallos_reportados",
    "FECHA DE FALLOS REPORTADOS": "fecha_fallos_reportados",
    "FECHA DE FALLOS CORREGIDOS": "fecha_fallos_corregidos",
    "CODIGO CHG": "codigo_chg",
    "FECHA SOLICITUD CHG": "fecha_solicitud_chg",
    "HORA DE SOLICITUD CHG": "hora_solicitud_chg",
    "FECHA EJECUCIÓN CHG": "fecha_ejecucion_chg",
    "HORA DE EJECUCION CHG": "hora_ejecucion_chg",
    "DÍAS QA": "dias_qa",
    "DÍAS INFRA/PLAT": "dias_infra_plat",
    "TOTAL DÍAS": "total_dias",
    "ESTADO QA": "estado_qa",
    "NOTAS DE SEGUIMIENTO": "notas_seguimiento",
    "OBSERVADO POR SOX": "observado_por_sox",
    "FECHA COMUNICADO": "fecha_comunicado",
    "FECHA REMEDIACIÓN": "fecha_remediacion",
    "TOTAL DÍAS REMEDIACIÓN": "total_dias_remediacion",
    "TOTAL DÍAS OBS": "total_dias_obs",
    "ROLLBACK": "rollback",
    "COMENTARIOS": "comentarios",
    "PROBLEMA EN DESPLIEGUE PROD": "problema_en_despliegue_prod",
    "TIPO DE PROBLEMA": "tipo_problema",
    "FECHA DE REPORTE": "fecha_reporte",
    "FECHA DE ATENCIÓN": "fecha_atencion",
    "TOTAL DÍAS CORRECCION PROD": "total_dias_correccion_prod",
    "TICKETS ATENDIDOS DENTRO DEL SLA": "tickets_dentro_sla",
    "CANT. C.P. QA": "cant_cp_qa",
    "CANT. CP. DEV": "cant_cp_dev",
    "FECHA REGULARIZACIÓN PPE": "fecha_regularizacion_ppe",
    "TOTAL DÍAS REGULARIZACIÓN": "total_dias_regularizacion",
    "MOTIVO PPE": "motivo_ppe",
}

ORDERED_STAGING_COLS = list(COLUMN_MAP.values()) + ["_file_name"]

# ========= LECTOR con fallbacks =========
def read_excel_any(path: Path, sheet: str) -> pd.DataFrame:
    src = Path(path)
    if not src.exists():
        raise FileNotFoundError(f"No se encuentra el archivo: {src}")
    try:
        LOCAL_COPY.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, LOCAL_COPY)
        print(f"📥 Copiado a {LOCAL_COPY}")
        p = LOCAL_COPY
    except Exception as e:
        print(f"⚠️ No se pudo copiar a {LOCAL_COPY}: {e}")
        p = src

    for engine_name in ["openpyxl", "xlrd", "pyxlsb"]:
        try:
            return pd.read_excel(p, sheet_name=sheet, dtype=str, engine=engine_name)
        except Exception as e:
            print(f"⚠️ {engine_name} falló: {e}")
    raise RuntimeError("No se pudo leer el Excel con ninguno de los motores (xlsx/xls/xlsb).")

# ========= MAIN =========
def main():
    df = read_excel_any(Path(EXCEL_PATH), EXCEL_SHEET)
    df.columns = [c.strip() for c in df.columns]

    present = [c for c in COLUMN_MAP if c in df.columns]
    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        print("⚠️ Faltan columnas (revisa el nombre exacto de la hoja y encabezados):", missing)

    df = df[present].copy().rename(columns=COLUMN_MAP)
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    df["_file_name"] = Path(EXCEL_PATH).name
    for col in ORDERED_STAGING_COLS:
        if col not in df.columns:
            df[col] = None
    df = df[ORDERED_STAGING_COLS]

    with engine.begin() as conn:
        df.to_sql(
            name="stg_tickets_diarios",
            schema="dbo",
            con=conn,
            if_exists="append",
            index=False,
            method="multi",
            chunksize=40,  # <= evita error de 2100 parámetros
        )
        print(f"✅ Insertadas {len(df)} filas en dbo.stg_tickets_diarios")

        conn.execute(
            text("EXEC dbo.sp_etl_tickets_diarios_to_final @source_file = :fname"),
            {"fname": Path(EXCEL_PATH).name}
        )
        print("✅ ETL ejecutado (staging -> qa_liu_2025)")

if __name__ == "__main__":
    main()
