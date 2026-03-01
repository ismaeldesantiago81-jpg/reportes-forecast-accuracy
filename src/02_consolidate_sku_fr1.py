from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" if (ROOT / "data").exists() else ROOT / "Data"
OUTPUT_DIR = ROOT / "output" if (ROOT / "output").exists() else ROOT / "Output"

import pandas as pd
import numpy as np

EXCEL_FILE = DATA_DIR / "P02 IBERIA.xlsx"
SHEET_NAME = "Todos"

COL_FR1 = "FR1"
COL_SKU = "Artículo ID"
COL_SALES = "A (Vts)"
COL_FCST = "F (DCH)"
COL_PERIOD = "Mes Año"

OUT_FILE = OUTPUT_DIR / "consolidated_sku_fr1_last_month.xlsx"

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

def main():
    if not EXCEL_FILE.exists():
        print(f"ERROR: File not found: {EXCEL_FILE.resolve()}")
        return

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    df[COL_SALES] = pd.to_numeric(df[COL_SALES], errors="coerce").fillna(0)
    df[COL_FCST] = pd.to_numeric(df[COL_FCST], errors="coerce").fillna(0)

    periods = df[COL_PERIOD].dropna().astype(str).unique().tolist()
    periods_sorted = sorted(periods, key=period_to_key)
    last_period = periods_sorted[-1] if periods_sorted else None

    if not last_period:
        print(f"ERROR: Could not detect last period from column: {COL_PERIOD}")
        return

    df = df[df[COL_PERIOD].astype(str) == str(last_period)].copy()

    df = df[~((df[COL_SALES] == 0) & (df[COL_FCST] == 0))].copy()

    g = df.groupby([COL_FR1, COL_SKU], dropna=False, as_index=False).agg(
        Ventas_SKU=(COL_SALES, "sum"),
        DCH_SKU=(COL_FCST, "sum"),
    )

    g["AbsError_SKU"] = (g["DCH_SKU"] - g["Ventas_SKU"]).abs()
    g["Over_SKU"] = np.maximum(g["DCH_SKU"] - g["Ventas_SKU"], 0)
    g["Gap_SKU"] = np.maximum(g["Ventas_SKU"] - g["DCH_SKU"], 0)
    g["NoDemand_SKU"] = np.where(g["Ventas_SKU"] == 0, g["DCH_SKU"], 0)

    OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    g.to_excel(OUT_FILE, index=False)

    print(f"Last period detected: {last_period}")
    print(f"Rows (filtered): {len(df)}")
    print(f"Rows (SKU+FR1 consolidated): {len(g)}")
    print(f"Saved: {OUT_FILE.as_posix()}")

if __name__ == "__main__":
    main()
