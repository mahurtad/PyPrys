import pandas as pd
from sqlalchemy import create_engine, text

# ========= CONFIG =========
SQL_SERVER   = r"localhost\SQLEXPRESS"
SQL_DATABASE = "qa_liu_2025"
SQL_USERNAME = "mahurtad"
SQL_PASSWORD = "NICAMINICA"

# ========= CONEXIÓN =========
def make_engine(driver_name: str):
    odbc = (
        f"Driver={{{driver_name}}};"
        f"Server={SQL_SERVER};Database={SQL_DATABASE};UID={SQL_USERNAME};PWD={SQL_PASSWORD};"
        "TrustServerCertificate=yes;"
    )
    return create_engine(f"mssql+pyodbc:///?odbc_connect={odbc}")

engine = None
for drv in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
    try:
        engine = make_engine(drv)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        print(f"✅ Conectado con: {drv}")
        break
    except Exception as e:
        print(f"⚠️ No se pudo con {drv}: {e}")
if engine is None:
    raise RuntimeError("No fue posible conectarse con ODBC 18 ni 17.")

# ========= CONSULTA =========
with engine.connect() as conn:
    # obtenemos todas las columnas de qa_liu_2025
    cols = pd.read_sql("""
        SELECT COLUMN_NAME
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = 'qa_liu_2025'
        ORDER BY ORDINAL_POSITION
    """, conn)["COLUMN_NAME"].tolist()

    results = {}
    for c in cols:
        q = f"SELECT COUNT(*) AS cnt FROM dbo.qa_liu_2025 WHERE [{c}] IS NOT NULL"
        cnt = conn.execute(text(q)).scalar()
        results[c] = cnt

# ========= MOSTRAR RESULTADOS =========
df_counts = pd.DataFrame(list(results.items()), columns=["columna", "no_nulos"]).sort_values("no_nulos", ascending=False)
print(df_counts)

# (Opcional: exportar a Excel para revisar cómodo)
df_counts.to_excel("C:/temp/qa_liu_2025_column_counts.xlsx", index=False)
