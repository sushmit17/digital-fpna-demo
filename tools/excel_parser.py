"""Parse uploaded unit submission Excel files into a standard dict structure."""
import pandas as pd
from pathlib import Path

def parse_unit_file(filepath: str) -> dict:
    """
    Returns:
        {
          "unit": str,
          "lines": {line_name: {lbe, prior_lbe, budget, prior_year}},
          "products": [  # only for Sales units (have Product Detail tab)
              {"name": str, "lbe_vol": float, "bgt_vol": float, "py_vol": float}
          ]
        }
    """
    xl = pd.ExcelFile(filepath)

    # ── P&L summary (LBE FY2026 tab) ─────────────────────────────────────────
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

    # ── Product detail (only Sales units have this tab) ───────────────────────
    products = []
    if "Product Detail" in xl.sheet_names:
        try:
            # Headers are on rows 5+6; data from row 7
            # Cols: A=Product, B=LBE vol, C=Prior LBE vol, D=Bgt vol, E=PY vol
            pdf = pd.read_excel(filepath, sheet_name="Product Detail",
                                header=None, skiprows=6)
            for _, row in pdf.iterrows():
                name = str(row[0]).strip() if pd.notna(row[0]) else ""
                if not name or name.startswith("Total") or name.startswith("Note"):
                    continue
                products.append({
                    "name":     name,
                    "lbe_vol":  _safe_float(row[1]),
                    "plbe_vol": _safe_float(row[2]),
                    "bgt_vol":  _safe_float(row[3]),
                    "py_vol":   _safe_float(row[4]),
                })
        except Exception:
            pass

    unit = Path(filepath).stem.replace("_LBE_FY2026", "").replace("_", " ").title()
    return {"unit": unit, "lines": lines, "products": products}


def _safe_float(val):
    try:
        f = float(val)
        return round(f, 1)
    except (TypeError, ValueError):
        return 0.0
