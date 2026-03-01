from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" if (ROOT / "data").exists() else ROOT / "Data"
OUTPUT_DIR = ROOT / "output" if (ROOT / "output").exists() else ROOT / "Output"

import pandas as pd
import numpy as np

# Inputs already produced by previous steps
CONSOLIDATED = OUTPUT_DIR / "consolidated_sku_fr1_last_month.xlsx"
TOP10_FILE = OUTPUT_DIR / "top10_abserror.xlsx"
BU_FILE = OUTPUT_DIR / "table_bu_by_fr1.xlsx"
SUBPLAT_FILE = OUTPUT_DIR / "table_subplatform_by_fr1.xlsx"

# Raw file used only to detect last period label for the filename + methodology header
RAW_FILE = DATA_DIR / "P02 IBERIA.xlsx"
SHEET_NAME = "Todos"
COL_PERIOD = "Mes Año"

MONTH_MAP = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12
}

def period_to_key(s: str) -> int:
    if not isinstance(s, str):
        return -1
    parts = s.strip().lower().split()
    if len(parts) != 2:
        return -1
    mon_txt, year_txt = parts
    mon = MONTH_MAP.get(mon_txt[:3], None)
    try:
        year = int(year_txt)
    except ValueError:
        return -1
    if mon is None:
        return -1
    return year * 100 + mon

def detect_last_period() -> str:
    raw = pd.read_excel(RAW_FILE, sheet_name=SHEET_NAME)
    periods = raw[COL_PERIOD].dropna().astype(str).unique().tolist()
    if not periods:
        return "unknown_period"
    return sorted(periods, key=period_to_key)[-1]

def build_global_kpis(g: pd.DataFrame) -> pd.DataFrame:
    acc_universe = g[g["Ventas_SKU"] > 0].copy()

    ventas_total = acc_universe["Ventas_SKU"].sum()
    dch_total = acc_universe["DCH_SKU"].sum()
    abserror_total = acc_universe["AbsError_SKU"].sum()

    accuracy = max(0.0, 1.0 - (abserror_total / ventas_total)) if ventas_total > 0 else np.nan
    bias = ((dch_total - ventas_total) / ventas_total) if ventas_total > 0 else np.nan

    over_total = acc_universe["Over_SKU"].sum()
    gap_total = acc_universe["Gap_SKU"].sum()
    nodemand_total = g.loc[g["Ventas_SKU"] == 0, "DCH_SKU"].sum()

    df = pd.DataFrame([{
        "Ventas_total_u": int(round(ventas_total)),
        "DCH_total_u": int(round(dch_total)),
        "AbsError_total_u": int(round(abserror_total)),
        "Over_Forecast_total_u": int(round(over_total)),
        "Forecast_Gap_total_u": int(round(gap_total)),
        "No_Demand_total_u": int(round(nodemand_total)),
        "Accuracy": accuracy,
        "Accuracy_%": accuracy * 100 if not pd.isna(accuracy) else np.nan,
        "Bias": bias,
        "Bias_%": bias * 100 if not pd.isna(bias) else np.nan,
    }])
    return df

def build_fr1_summary(g: pd.DataFrame) -> pd.DataFrame:
    # Accuracy/Bias governed on Ventas>0 universe
    acc = g[g["Ventas_SKU"] > 0].groupby("FR1", dropna=False, as_index=False).agg(
        Ventas=("Ventas_SKU", "sum"),
        DCH=("DCH_SKU", "sum"),
        AbsError=("AbsError_SKU", "sum"),
        Over_Forecast=("Over_SKU", "sum"),
        Forecast_Gap=("Gap_SKU", "sum"),
    )
    nod = g[g["Ventas_SKU"] == 0].groupby("FR1", dropna=False, as_index=False).agg(
        No_Demand=("DCH_SKU", "sum")
    )
    out = acc.merge(nod, on="FR1", how="left")
    out["No_Demand"] = out["No_Demand"].fillna(0)

    out["Accuracy"] = np.maximum(0.0, 1.0 - (out["AbsError"] / out["Ventas"]))
    out["Bias"] = (out["DCH"] - out["Ventas"]) / out["Ventas"]

    for c in ["Ventas", "DCH", "AbsError", "Over_Forecast", "Forecast_Gap", "No_Demand"]:
        out[c] = out[c].round().astype(int)

    out = out.sort_values(["AbsError"], ascending=False).reset_index(drop=True)
    return out

def methodology_text(period: str) -> pd.DataFrame:
    lines = [
        f"Periodo: {period} (último mes detectado en el Excel)",
        "",
        "Tratamiento de datos faltantes:",
        "- Ventas NaN = 0",
        "- DCH NaN = 0",
        "- Filas con Ventas=0 y DCH=0 → excluidas del universo",
        "",
        "Consolidación (Paso 2):",
        "- Nivel: FR1 (KAM) + SKU (Artículo ID)",
        "- Ventas_SKU = Σ Ventas",
        "- DCH_SKU = Σ DCH",
        "- AbsError_SKU = |DCH_SKU − Ventas_SKU|",
        "- Over-Forecast_SKU = max(DCH_SKU − Ventas_SKU, 0)",
        "- Forecast Gap_SKU = max(Ventas_SKU − DCH_SKU, 0)",
        "- No-Demand_SKU = DCH_SKU donde Ventas_SKU = 0 (fuera de Accuracy)",
        "",
        "KPIs gobernados (Paso 3) con SKUs Ventas>0:",
        "- Accuracy = MAX(0, 1 − Σ|DCH−Ventas| / ΣVentas)",
        "- Bias = (ΣDCH − ΣVentas) / ΣVentas",
        "- Over-Forecast (u) = Σ max(DCH−Ventas, 0)",
        "- Forecast Gap (u) = Σ max(Ventas−DCH, 0)",
        "- No-Demand (u) separado",
        "",
        "Regla crítica:",
        "- AbsError se calcula a nivel SKU y se agrega después (el error no se compensa).",
        "",
        "Coherencia (OK/KO):",
        "- Σ Ventas (FR1) = Σ Ventas (Subplataforma)",
        "- Σ DCH (FR1) = Σ DCH (Subplataforma)",
    ]
    return pd.DataFrame({"Methodology": lines})

def main():
    # Basic existence checks
    needed = [CONSOLIDATED, BU_FILE, SUBPLAT_FILE]
    missing = [p.as_posix() for p in needed if not p.exists()]
    if missing:
        print("[ERROR] RESULT: KO (Missing required files)")
        for m in missing:
            print(" -", m)
        return

    period = detect_last_period()
    out_name = f"REPORT_ForecastAccuracy_{period.replace(' ', '_')}.xlsx"
    OUT = OUTPUT_DIR / out_name

    g = pd.read_excel(CONSOLIDATED)

    global_kpis = build_global_kpis(g)
    fr1_summary = build_fr1_summary(g)
    bu = pd.read_excel(BU_FILE)
    subplat = pd.read_excel(SUBPLAT_FILE)

    top10 = pd.read_excel(TOP10_FILE) if TOP10_FILE.exists() else pd.DataFrame()

    meth = methodology_text(period)

    with pd.ExcelWriter(OUT, engine="openpyxl") as writer:
        global_kpis.to_excel(writer, index=False, sheet_name="00_Global_KPIs")
        fr1_summary.to_excel(writer, index=False, sheet_name="01_FR1_Summary")
        bu.to_excel(writer, index=False, sheet_name="02_BU_by_FR1")
        subplat.to_excel(writer, index=False, sheet_name="03_Subplatform_by_FR1")
        if len(top10) > 0:
            top10.to_excel(writer, index=False, sheet_name="04_Top10_AbsError")
        else:
            pd.DataFrame({"Info": ["Top10 file not found: run src/04_top10_abserror.py"]}).to_excel(
                writer, index=False, sheet_name="04_Top10_AbsError"
            )
        meth.to_excel(writer, index=False, sheet_name="99_Methodology")

    print(f"[OK] RESULT: OK (Final report created) -> {OUT.as_posix()}")

if __name__ == "__main__":
    main()
