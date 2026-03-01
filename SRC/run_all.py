import subprocess
import sys
from pathlib import Path

PYTHON = sys.executable  # ensures we use the current interpreter (your .venv)

STEPS = [
    ("02 Consolidate SKU+FR1 (last month)", [PYTHON, "src/02_consolidate_sku_fr1.py"]),
    ("04 Top10 AbsError",                   [PYTHON, "src/04_top10_abserror.py"]),
    ("05 Coherence check",                  [PYTHON, "src/05_coherence_check.py"]),
    ("06 Table Subplatform by FR1",         [PYTHON, "src/06_table_subplatform_by_fr1.py"]),
    ("07 Table BU by FR1",                  [PYTHON, "src/07_table_bu_by_fr1.py"]),
    ("03 Global KPIs (console)",            [PYTHON, "src/03_kpis_global.py"]),
    ("08 Build final report",               [PYTHON, "src/08_build_final_report.py"]),
]

def run_step(name, cmd):
    print(f"\n=== {name} ===")
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.stdout:
        print(result.stdout.rstrip())
    if result.stderr:
        print(result.stderr.rstrip(), file=sys.stderr)

    if result.returncode != 0:
        print(f"\n🔴 STOP: Step failed (non-zero exit): {name}")
        sys.exit(result.returncode)

    if "RESULT: KO" in (result.stdout or ""):
        print(f"\n🔴 STOP: KO detected in step: {name}")
        sys.exit(2)

def main():
    if not Path("src").exists():
        print("❌ Can't find src/ folder. Run this from the project root (Forecast_accuracy).")
        sys.exit(1)

    # quick debug: show which python we're using
    print(f"🐍 Using interpreter: {PYTHON}")

    for name, cmd in STEPS:
        run_step(name, cmd)

    print("\n🟢 ALL DONE: Report pipeline finished successfully.")

if __name__ == "__main__":
    main()
