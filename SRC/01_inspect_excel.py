from pathlib import Path
import pandas as pd

# 1) Put your Excel file inside the /data folder and write its name here
EXCEL_FILE = Path("data") / "P02 IBERIA.xlsx"

def main():
    if not EXCEL_FILE.exists():
        print(f"❌ File not found: {EXCEL_FILE.resolve()}")
        print("Put your Excel inside the data/ folder and update EXCEL_FILE.")
        return

    # List sheet names
    xls = pd.ExcelFile(EXCEL_FILE)
    print("📄 Sheets found:")
    for s in xls.sheet_names:
        print(" -", s)

    # Read first sheet by default (we can change this once we know the right one)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)

    print(f"\n✅ Loaded sheet: {sheet}")
    print("\n🧾 Columns:")
    print(list(df.columns))

    print("\n👀 First 5 rows:")
    print(df.head(5).to_string(index=False))

if __name__ == "__main__":
    main()