import argparse
import sys
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt

# -----------------------------
# Helpers
# -----------------------------
DISPLAY_WIDTH = 160

def exists(col, df):
    return col in df.columns

def to_datetime(series):
    return pd.to_datetime(series, errors="coerce", dayfirst=False, infer_datetime_format=True)

def to_numeric(series):
    return pd.to_numeric(series, errors="coerce")

def safe_pct(n, d):
    return (n / d * 100.0) if (d and d != 0) else 0.0

def print_header(title):
    print("\n" + "=" * DISPLAY_WIDTH)
    print(title)
    print("=" * DISPLAY_WIDTH)

def save_fig(plt_obj, outdir: Path, name: str):
    out = outdir / f"{name}.png"
    plt_obj.tight_layout()
    plt_obj.savefig(out, dpi=140, bbox_inches="tight")
    plt_obj.close()
    return out

# -----------------------------
# Main
# -----------------------------
def main():
    parser = argparse.ArgumentParser(description="Console metrics for 'Tickets Diarios'")
    parser.add_argument("--path", type=str, default=r"C:\Users\manue\OneDrive - EduCorpPERU\Calidad de Software - Certificaciones\Gestión Demanda Certificaciones LIUv1.xlsx",
                        help="Ruta al Excel con la hoja 'Tickets Diarios'")
    parser.add_argument("--sheet", type=str, default="Tickets Diarios", help="Nombre de la hoja")
    parser.add_argument("--outdir", type=str, default=str(Path.cwd()), help="Directorio de salida para TXT/CSV/PNG")
    args = parser.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    # Leer Excel
    try:
        df = pd.read_excel(args.path, sheet_name=args.sheet)
    except Exception as e:
        print("ERROR al leer el Excel:", e)
        print("Sugerencia: verifique la ruta --path y que la hoja exista.")
        sys.exit(1)

    # Normalizar headers (quitar espacios finales)
    df.columns = [str(c).strip() for c in df.columns]

    # Intentar parsear fechas relevantes si existen
    date_cols = [
        "FECHA RECEPCION COORDINADOR QA",
        "FECHA ASIGNACIÓN",
        "FECHA DE FALLOS REPORTADOS",
        "FECHA DE FALLOS CORREGIDOS",
        "FECHA SOLICITUD CHG",
        "FECHA EJECUCIÓN CHG",
        "FECHA COMUNICADO",
        "FECHA REMEDIACIÓN",
        "FECHA DE REPORTE",
        "FECHA DE ATENCIÓN",
    ]
    for c in date_cols:
        if exists(c, df):
            df[c] = to_datetime(df[c])

    # Intentar convertir numéricos
    num_cols = [
        "CANT. FALLOS REPORTADOS",
        "DÍAS QA",
        "DÍAS INFRA/PLAT",
        "TOTAL DÍAS",
        "TOTAL DÍAS REMEDIACIÓN",
        "TOTAL DÍAS OBS",
        "TOTAL DÍAS CORRECCION PROD",
        "TICKETS ATENDIDOS DENTRO DEL SLA",
    ]
    for c in num_cols:
        if exists(c, df):
            df[c] = to_numeric(df[c])

    # Derivados de fechas
    # Lead time: Recepción -> Asignación
    if exists("FECHA RECEPCION COORDINADOR QA", df) and exists("FECHA ASIGNACIÓN", df):
        df["LEAD_DIAS_QA_ASIGNACION"] = (df["FECHA ASIGNACIÓN"] - df["FECHA RECEPCION COORDINADOR QA"]).dt.total_seconds() / 86400.0

    # Reporte -> Atención en Prod
    if exists("FECHA DE REPORTE", df) and exists("FECHA DE ATENCIÓN", df):
        df["LEAD_DIAS_REPORTE_ATENCION"] = (df["FECHA DE ATENCIÓN"] - df["FECHA DE REPORTE"]).dt.total_seconds() / 86400.0

    # CHG: Solicitud -> Ejecución
    if exists("FECHA SOLICITUD CHG", df) and exists("FECHA EJECUCIÓN CHG", df):
        df["LEAD_DIAS_CHG"] = (df["FECHA EJECUCIÓN CHG"] - df["FECHA SOLICITUD CHG"]).dt.total_seconds() / 86400.0

    # Remediación: Comunicado -> Remediación
    if exists("FECHA COMUNICADO", df) and exists("FECHA REMEDIACIÓN", df):
        df["LEAD_DIAS_REMEDIACION"] = (df["FECHA REMEDIACIÓN"] - df["FECHA COMUNICADO"]).dt.total_seconds() / 86400.0

    # Fecha base para series mensuales: usar FECHA DE REPORTE si existe; si no, Recepción
    if exists("FECHA DE REPORTE", df):
        df["_FECHA_BASE"] = df["FECHA DE REPORTE"]
    elif exists("FECHA RECEPCION COORDINADOR QA", df):
        df["_FECHA_BASE"] = df["FECHA RECEPCION COORDINADOR QA"]
    else:
        df["_FECHA_BASE"] = pd.NaT

    if df["_FECHA_BASE"].notna().any():
        df["_YM"] = df["_FECHA_BASE"].dt.to_period("M").astype(str)

    # -----------------------------
    # Consola: Métricas principales
    # -----------------------------
    total_tickets = len(df)
    print_header("MÉTRICAS GENERALES")
    print(f"Total de tickets: {total_tickets}")

    # Por institución
    if exists("INSTITUCION", df):
        institucion_ct = df["INSTITUCION"].value_counts(dropna=False).rename_axis("INSTITUCION").reset_index(name="TICKETS")
        print_header("Tickets por INSTITUCIÓN")
        print(institucion_ct.to_string(index=False))
        institucion_ct.to_csv(Path(outdir, "tickets_por_institucion.csv"), index=False)

    # Por estado QA
    if exists("ESTADO QA", df):
        estado_ct = df["ESTADO QA"].value_counts(dropna=False).rename_axis("ESTADO QA").reset_index(name="TICKETS")
        print_header("Tickets por ESTADO QA")
        print(estado_ct.to_string(index=False))
        estado_ct.to_csv(Path(outdir, "tickets_por_estado_qa.csv"), index=False)

    # SLA
    sla_pct = None
    if exists("TICKETS ATENDIDOS DENTRO DEL SLA", df):
        dentro = (df["TICKETS ATENDIDOS DENTRO DEL SLA"] == 1).sum()
        sla_pct = safe_pct(dentro, total_tickets)
        print_header("SLA")
        print(f"Tickets dentro de SLA: {dentro} de {total_tickets} ({sla_pct:.2f}%)")

    # Problema despliegue / rollback
    if exists("PROBLEMA EN DESPLIEGUE PROD", df):
        prob_ct = df["PROBLEMA EN DESPLIEGUE PROD"].fillna("Sin dato").value_counts(dropna=False).rename_axis("PROBLEMA EN DESPLIEGUE PROD").reset_index(name="TICKETS")
        print_header("Problema en Despliegue Prod")
        print(prob_ct.to_string(index=False))
        prob_ct.to_csv(Path(outdir, "problema_despliegue_prod.csv"), index=False)

    if exists("ROLLBACK", df):
        rb_ct = df["ROLLBACK"].fillna("Sin dato").value_counts(dropna=False).rename_axis("ROLLBACK").reset_index(name="TICKETS")
        print_header("Rollback")
        print(rb_ct.to_string(index=False))
        rb_ct.to_csv(Path(outdir, "rollback.csv"), index=False)

    # Tipo de cambio / Impacto / Riesgo y otros categóricos clave
    for col in ["TIPO CAMBIO", "IMPACTO", "RIESGO", "AREA", "AREA DESARROLLO", "DESARROLLADOR", "SOLICITANTE", "APLICATIVO/BD", "TIPO DE PROBLEMA"]:
        if exists(col, df):
            ct = df[col].fillna("Sin dato").value_counts(dropna=False).head(30).rename_axis(col).reset_index(name="TICKETS")
            print_header(f"Top categorías: {col} (Top 30)")
            print(ct.to_string(index=False))
            ct.to_csv(Path(outdir, f"top_{col.replace('/', '_').replace(' ', '_').lower()}.csv"), index=False)

    # Fallos reportados
    if exists("CANT. FALLOS REPORTADOS", df):
        total_fallos = df["CANT. FALLOS REPORTADOS"].sum(skipna=True)
        avg_fallos = df["CANT. FALLOS REPORTADOS"].mean(skipna=True)
        print_header("Fallos reportados")
        print(f"Total fallos reportados: {int(total_fallos) if pd.notna(total_fallos) else 0}")
        print(f"Promedio de fallos por ticket: {avg_fallos:.2f}" if pd.notna(avg_fallos) else "Promedio no disponible")

    # Métricas de días (promedio/percentiles)
    def report_days(col):
        if exists(col, df):
            series = to_numeric(df[col]).dropna()
            if len(series):
                p50 = series.quantile(0.50)
                p75 = series.quantile(0.75)
                p90 = series.quantile(0.90)
                p95 = series.quantile(0.95)
                print_header(f"Distribución de días: {col}")
                print(f"Media: {series.mean():.2f} | Mediana (p50): {p50:.2f} | p75: {p75:.2f} | p90: {p90:.2f} | p95: {p95:.2f} | Máx: {series.max():.2f}")
            else:
                print_header(f"Distribución de días: {col}")
                print("Sin datos numéricos")

    for col in [
        "DÍAS QA", "DÍAS INFRA/PLAT", "TOTAL DÍAS", "TOTAL DÍAS REMEDIACIÓN", "TOTAL DÍAS OBS",
        "TOTAL DÍAS CORRECCION PROD", "LEAD_DIAS_QA_ASIGNACION", "LEAD_DIAS_REPORTE_ATENCION",
        "LEAD_DIAS_CHG", "LEAD_DIAS_REMEDIACION"
    ]:
        report_days(col)

    # Series mensuales
    if "_YM" in df.columns:
        monthly = df.groupby("_YM").size().rename("TICKETS").reset_index()
        monthly = monthly.sort_values("_YM")
        print_header("Tendencia mensual (tickets)")
        print(monthly.to_string(index=False))
        monthly.to_csv(Path(outdir, "tendencia_mensual_tickets.csv"), index=False)

        # Gráfico simple
        plt.figure(figsize=(10, 4))
        plt.plot(monthly["_YM"], monthly["TICKETS"], marker="o")
        plt.xticks(rotation=45, ha="right")
        plt.title("Tickets por mes")
        plt.xlabel("Mes")
        plt.ylabel("Tickets")
        save_fig(plt, outdir, "tendencia_mensual_tickets")

        # SLA mensual si es posible
        if exists("TICKETS ATENDIDOS DENTRO DEL SLA", df):
            df["_DENTRO_SLA"] = (df["TICKETS ATENDIDOS DENTRO DEL SLA"] == 1).astype(int)
            sla_month = df.groupby("_YM")["_DENTRO_SLA"].sum().rename("DENTRO_SLA").reset_index()
            sla_month = sla_month.merge(monthly, on="_YM", how="left")
            sla_month["SLA_%"] = np.where(sla_month["TICKETS"] > 0, sla_month["DENTRO_SLA"] / sla_month["TICKETS"] * 100.0, np.nan)
            print_header("Tendencia mensual SLA %")
            print(sla_month[["_YM", "SLA_%"]].to_string(index=False, float_format=lambda x: f"{x:.2f}"))
            sla_month.to_csv(Path(outdir, "tendencia_mensual_sla.csv"), index=False)

            plt.figure(figsize=(10, 4))
            plt.plot(sla_month["_YM"], sla_month["SLA_%"], marker="o")
            plt.xticks(rotation=45, ha="right")
            plt.title("SLA % por mes")
            plt.xlabel("Mes")
            plt.ylabel("SLA %")
            save_fig(plt, outdir, "tendencia_mensual_sla")

    # Guardar resumen TXT
    txt_lines = []
    txt_lines.append(f"Total tickets: {total_tickets}")
    if sla_pct is not None:
        txt_lines.append(f"SLA global: {sla_pct:.2f}%")
    txt_content = "\n".join(txt_lines)
    (outdir / "metrics_summary.txt").write_text(txt_content, encoding="utf-8")

    print_header("Archivos generados")
    print(f"- metrics_summary.txt")
    print(f"- CSVs (varios): en {outdir}")
    print(f"- PNGs (si hay series mensuales): en {outdir}")
    print("\n✅ Listo. Métricas impresas arriba y archivos exportados.")

if __name__ == "__main__":
    pd.set_option("display.width", DISPLAY_WIDTH)
    pd.set_option("display.max_columns", 50)
    main()
