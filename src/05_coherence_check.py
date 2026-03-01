from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" if (ROOT / "data").exists() else ROOT / "Data"
OUTPUT_DIR = ROOT / "output" if (ROOT / "output").exists() else ROOT / "Output"

import pandas as pd
import numpy as np

EXCEL_FILE = DATA_DIR / "P02 IBERIA.xlsx"
SHEET_NAME = "Todos"

COL_GROUP = "FR1"
COL_SUBPLAT = "Subplataforma"
COL_SALES = "A (Vts)"
COL_FCST = "F (DCH)"
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

def main():
    if not EXCEL_FILE.exists():
        print(f"[ERROR] File not found: {EXCEL_FILE.resolve()}")
        return

    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

    # NaN -> 0 (governed rule)
    df[COL_SALES] = pd.to_numeric(df[COL_SALES], errors="coerce").fillna(0)
    df[COL_FCST] = pd.to_numeric(df[COL_FCST], errors="coerce").fillna(0)

    # Detect last period
    periods = df[COL_PERIOD].dropna().astype(str).unique().tolist()
    periods_sorted = sorted(periods, key=period_to_key)
    last_period = periods_sorted[-1] if periods_sorted else None
    if not last_period:
        print(f"[ERROR] Could not detect last period from {COL_PERIOD}")
        return

    # Filter to last period
    df = df[df[COL_PERIOD].astype(str) == str(last_period)].copy()

    # Universe rule: exclude rows with Sales=0 and Forecast=0
    df = df[~((df[COL_SALES] == 0) & (df[COL_FCST] == 0))].copy()

    # Totals computed two ways (must match)
    total_sales = df[COL_SALES].sum()
    total_fcst = df[COL_FCST].sum()

    by_group = df.groupby(COL_GROUP, dropna=False)[[COL_SALES, COL_FCST]].sum()
    by_subplat = df.groupby(COL_SUBPLAT, dropna=False)[[COL_SALES, COL_FCST]].sum()

    sales_group = by_group[COL_SALES].sum()
    fcst_group = by_group[COL_FCST].sum()

    sales_subplat = by_subplat[COL_SALES].sum()
    fcst_subplat = by_subplat[COL_FCST].sum()

    # Use exact comparison with a tiny tolerance (Excel floats can be messy)
    tol = 1e-9
    ok_sales = (abs(total_sales - sales_group) < tol) and (abs(total_sales - sales_subplat) < tol)
    ok_fcst = (abs(total_fcst - fcst_group) < tol) and (abs(total_fcst - fcst_subplat) < tol)

    print(f"[OK] Period: {last_period}")
    print(f"Ventas total (universe): {total_sales:.6f}")
    print(f"DCH total (universe):    {total_fcst:.6f}")

    print("\n[INFO] Coherence check:")
    print(f"Σ Ventas (FR1)          = {sales_group:.6f}")
    print(f"Σ Ventas (Subplataforma)= {sales_subplat:.6f}")
    print(f"Σ DCH (FR1)             = {fcst_group:.6f}")
    print(f"Σ DCH (Subplataforma)   = {fcst_subplat:.6f}")

    if ok_sales and ok_fcst:
        print("\n[OK] RESULT: OK (Reporte publicable en este control)")
    else:
        print("\n[ERROR] RESULT: KO (Totales no coherentes, NO publicable)")
        if not ok_sales:
            print(" - KO en Ventas")
        if not ok_fcst:
            print(" - KO en DCH")

if __name__ == '__main__':
    main()
