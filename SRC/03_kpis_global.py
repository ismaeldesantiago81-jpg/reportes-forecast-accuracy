from pathlib import Path
import pandas as pd
import numpy as np

IN_FILE = Path("output") / "consolidated_sku_fr1_last_month.xlsx"

def main():
    if not IN_FILE.exists():
        print(f"❌ Missing input: {IN_FILE.as_posix()}")
        print("Run src/02_consolidate_sku_fr1.py first.")
        return

    g = pd.read_excel(IN_FILE)

    # Universe for Accuracy/Bias: only SKUs with Ventas>0
    acc_universe = g[g["Ventas_SKU"] > 0].copy()

    ventas_total = acc_universe["Ventas_SKU"].sum()
    dch_total = acc_universe["DCH_SKU"].sum()
    abserror_total = acc_universe["AbsError_SKU"].sum()

    # Governed KPIs
    accuracy = max(0.0, 1.0 - (abserror_total / ventas_total)) if ventas_total > 0 else np.nan
    bias = ((dch_total - ventas_total) / ventas_total) if ventas_total > 0 else np.nan

    over_total = acc_universe["Over_SKU"].sum()
    gap_total = acc_universe["Gap_SKU"].sum()

    # No-Demand: DCH where Ventas=0 (excluded from Accuracy)
    nodemand_total = g.loc[g["Ventas_SKU"] == 0, "DCH_SKU"].sum()

    # Round after aggregation (integers for units; % keep a few decimals for display)
    print("📌 GLOBAL KPIs (Governed)")
    print(f"Ventas_total (u): {int(round(ventas_total))}")
    print(f"DCH_total (u): {int(round(dch_total))}")
    print(f"AbsError_total (u): {int(round(abserror_total))}")
    print(f"Over-Forecast_total (u): {int(round(over_total))}")
    print(f"Forecast Gap_total (u): {int(round(gap_total))}")
    print(f"No-Demand_total (u): {int(round(nodemand_total))}")

    if pd.isna(accuracy):
        print("Accuracy: No dispongo de ese dato (Ventas_total=0 en el universo).")
    else:
        print(f"Accuracy: {accuracy:.4f}  ({accuracy*100:.2f}%)")

    if pd.isna(bias):
        print("Bias: No dispongo de ese dato (Ventas_total=0 en el universo).")
    else:
        print(f"Bias: {bias:.4f}  ({bias*100:.2f}%)")

if __name__ == "__main__":
    main()
