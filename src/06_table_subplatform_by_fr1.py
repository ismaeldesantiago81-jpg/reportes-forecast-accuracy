from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" if (ROOT / "data").exists() else ROOT / "Data"
OUTPUT_DIR = ROOT / "output" if (ROOT / "output").exists() else ROOT / "Output"

import pandas as pd
import numpy as np

IN_FILE = OUTPUT_DIR / "consolidated_sku_fr1_last_month.xlsx"
OUT_FILE = OUTPUT_DIR / "table_subplatform_by_fr1.xlsx"

EXCEL_FILE = DATA_DIR / "P02 IBERIA.xlsx"
SHEET_NAME = "Todos"

COL_PERIOD = "Mes Año"
COL_FR1 = "FR1"
COL_SKU = "Artículo ID"
COL_SUBPLAT = "Subplataforma"
COL_SALES = "A (Vts)"
COL_FCST = "F (DCH)"

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

def detect_last_period(df: pd.DataFrame) -> str | None:
    periods = df[COL_PERIOD].dropna().astype(str).unique().tolist()
    if not periods:
        return None
    return sorted(periods, key=period_to_key)[-1]

def main():
    if not IN_FILE.exists():
        print(f"[ERROR] Missing input: {IN_FILE.as_posix()}")
        print("Run src/02_consolidate_sku_fr1.py first.")
        return

    # 1) Load consolidated SKU+FR1 metrics (already has Ventas_SKU, DCH_SKU, AbsError_SKU, Over, Gap, NoDemand)
    g = pd.read_excel(IN_FILE)

    # 2) We need Subplataforma attached to each (FR1, SKU) from the raw file
    if not EXCEL_FILE.exists():
        print(f"[ERROR] Missing raw file: {EXCEL_FILE.as_posix()}")
        return

    raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    # Apply same universe/period rules as earlier so mapping is consistent
    raw[COL_SALES] = pd.to_numeric(raw[COL_SALES], errors="coerce").fillna(0)
    raw[COL_FCST] = pd.to_numeric(raw[COL_FCST], errors="coerce").fillna(0)

    last_period = detect_last_period(raw)
    if not last_period:
        print(f"[ERROR] Could not detect last period from {COL_PERIOD}")
        return

    raw = raw[raw[COL_PERIOD].astype(str) == str(last_period)].copy()
    raw = raw[~((raw[COL_SALES] == 0) & (raw[COL_FCST] == 0))].copy()

    # For each (FR1, SKU), take the literal Subplataforma.
    # If a (FR1, SKU) appears with multiple Subplataformas, we must flag it (traceability/coherence risk).
    map_df = (
        raw.groupby([COL_FR1, COL_SKU], dropna=False)[COL_SUBPLAT]
           .nunique(dropna=False)
           .reset_index(name="n_subplataformas")
    )

    conflicts = map_df[map_df["n_subplataformas"] > 1]
    if len(conflicts) > 0:
        print("[ERROR] RESULT: KO (SKU appears in multiple Subplataformas within same FR1 in the period)")
        print(conflicts.head(20).to_string(index=False))
        print("Fix the source mapping before publishing.")
        return

    mapping = (
        raw.groupby([COL_FR1, COL_SKU], dropna=False)[COL_SUBPLAT]
           .first()
           .reset_index()
    )

    # 3) Attach Subplataforma to consolidated table
    g2 = g.merge(mapping, left_on=["FR1", "Artículo ID"], right_on=[COL_FR1, COL_SKU], how="left")

    if g2[COL_SUBPLAT].isna().any():
        missing = g2[g2[COL_SUBPLAT].isna()][["FR1", "Artículo ID"]].head(20)
        print("[ERROR] RESULT: KO (Missing Subplataforma mapping for some (FR1, SKU))")
        print(missing.to_string(index=False))
        return

    # 4) Build table by (FR1, Subplataforma) using governed aggregation rules
    # Accuracy/Bias use only SKUs with Ventas_SKU>0 (exclude No-Demand)
    acc_universe = g2[g2["Ventas_SKU"] > 0].copy()

    agg = acc_universe.groupby(["FR1", COL_SUBPLAT], dropna=False, as_index=False).agg(
        Ventas=("Ventas_SKU", "sum"),
        DCH=("DCH_SKU", "sum"),
        AbsError=("AbsError_SKU", "sum"),
        Over_Forecast=("Over_SKU", "sum"),
        Forecast_Gap=("Gap_SKU", "sum"),
    )

    # No-Demand by (FR1, Subplataforma): from SKUs with Ventas=0
    nod = g2[g2["Ventas_SKU"] == 0].groupby(["FR1", COL_SUBPLAT], dropna=False, as_index=False).agg(
        No_Demand=("DCH_SKU", "sum")
    )

    table = agg.merge(nod, on=["FR1", COL_SUBPLAT], how="left")
    table["No_Demand"] = table["No_Demand"].fillna(0)

    # KPIs (after aggregation)
    table["Accuracy"] = np.maximum(0.0, 1.0 - (table["AbsError"] / table["Ventas"]))
    table["Bias"] = (table["DCH"] - table["Ventas"]) / table["Ventas"]

    # Rounding rules: units as integers after aggregation
    for c in ["Ventas", "DCH", "AbsError", "Over_Forecast", "Forecast_Gap", "No_Demand"]:
        table[c] = table[c].round().astype(int)

    # Keep Accuracy/Bias numeric but readable
    table = table.sort_values(["FR1", "AbsError"], ascending=[True, False]).reset_index(drop=True)

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    table.to_excel(OUT_FILE, index=False)

    print(f"[OK] Period: {last_period}")
    print(f"[INFO] Saved: {OUT_FILE.as_posix()}")
    print("\n[INFO] Sample (first 10 rows):")
    print(table.head(10).to_string(index=False))

    print("\n[OK] RESULT: OK (Table generated)")

if __name__ == "__main__":
    main()
