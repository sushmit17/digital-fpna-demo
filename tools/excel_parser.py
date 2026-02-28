"""Parse uploaded unit submission Excel files into a standard dict structure."""
import pandas as pd
from pathlib import Path

EXPECTED_COLS = ["P&L Line", "LBE FY2026", "Prior LBE", "Budget FY2026", "Prior Year (FY2025)"]

def parse_unit_file(filepath: str) -> dict:
    """
    Returns:
        {
          "unit": str,
          "lines": {line_name: {lbe, prior_lbe, budget, prior_year}}
        }
    """
    df = pd.read_excel(filepath, sheet_name="LBE FY2026", header=3)
    df.columns = ["line", "lbe", "prior_lbe", "budget", "prior_year"]
    df = df.dropna(subset=["line"])
    df = df[df["line"].astype(str).str.strip() != ""]
    df = df[~df["line"].astype(str).str.startswith("Note:")]

    lines = {}
    for _, row in df.iterrows():
        name = str(row["line"]).strip()
        lines[name] = {
            "lbe":        _safe_float(row["lbe"]),
            "prior_lbe":  _safe_float(row["prior_lbe"]),
            "budget":     _safe_float(row["budget"]),
            "prior_year": _safe_float(row["prior_year"]),
        }

    # Derive unit name from filename
    unit = Path(filepath).stem.replace("_LBE_FY2026", "").replace("_", " ").title()
    return {"unit": unit, "lines": lines}

def _safe_float(val):
    try:
        f = float(val)
        return round(f, 1)
    except (TypeError, ValueError):
        return 0.0
