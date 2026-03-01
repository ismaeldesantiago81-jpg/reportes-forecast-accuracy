from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data" if (ROOT / "data").exists() else ROOT / "Data"
OUTPUT_DIR = ROOT / "output" if (ROOT / "output").exists() else ROOT / "Output"

import pandas as pd

IN_FILE = OUTPUT_DIR / "consolidated_sku_fr1_last_month.xlsx"

def main():
    if not IN_FILE.exists():
        print(f"[ERROR] Missing input: {IN_FILE.as_posix()}")
        print("Run src/02_consolidate_sku_fr1.py first.")
        return

    g = pd.read_excel(IN_FILE)

    top10 = (
        g.sort_values(["AbsError_SKU"], ascending=False)
         .head(10)
         .copy()
    )

    # Integer display after final aggregation (here it's already aggregated)
    for c in ["Ventas_SKU", "DCH_SKU", "AbsError_SKU", "Over_SKU", "Gap_SKU", "NoDemand_SKU"]:
        top10[c] = top10[c].round().astype(int)

    print("[INFO] TOP 10 (AbsError_SKU) after SKU+FR1 consolidation")
    print(top10.to_string(index=False))

    out = OUTPUT_DIR / "top10_abserror.xlsx"
    top10.to_excel(out, index=False)
    print(f"\n[INFO] Saved: {out.as_posix()}")

if __name__ == "__main__":
    main()
