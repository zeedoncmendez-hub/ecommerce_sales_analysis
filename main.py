

Runs the four project scripts in order:
1. clean.py
2. analysis.py
3. visualization.py
4. create_excel_report.py

Run this from the project root:
    python main.py
"""

from pathlib import Path
import subprocess
import sys

PROJECT_DIR = Path(__file__).resolve().parent
SRC_DIR = PROJECT_DIR / "src"

PROGRAMS = [
    ("Data cleaning", SRC_DIR / "clean.py"),
    ("Exploratory analysis and KPIs", SRC_DIR / "analysis.py"),
    ("Matplotlib visualizations", SRC_DIR / "visualization.py"),
    ("Excel report generation", SRC_DIR / "excel_report.py"),
]


def run_program(name: str, script_path: Path) -> None:
    """Run one Python script and stop if it fails."""
    print("\n" + "=" * 70)
    print(f"STARTING: {name}")
    print(f"Script: {script_path}")
    print("=" * 70)

    if not script_path.exists():
        raise FileNotFoundError(
            f"Cannot find {script_path}\n"
            f"Check that {script_path.name} exists in your src folder."
        )

    result = subprocess.run(
        [sys.executable, str(script_path)],
        cwd=PROJECT_DIR,
        check=False,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"{name} failed with exit code {result.returncode}. "
            "The remaining programs were not run."
        )

    print(f"\nCOMPLETED: {name}")


def main() -> None:
    print("E-COMMERCE SALES ANALYSIS — MAIN PROGRAM")
    print(f"Project folder: {PROJECT_DIR}")

    try:
        for name, script_path in PROGRAMS:
            run_program(name, script_path)
    except (FileNotFoundError, RuntimeError) as error:
        print(f"\nERROR: {error}")
        sys.exit(1)

    print("\n" + "=" * 70)
    print("ALL FOUR PROGRAMS COMPLETED SUCCESSFULLY.")
    print("Check your Data/Cleaned, outputs/charts, and Excel report folders.")
    print("=" * 70)


if __name__ == "__main__":
    main()

