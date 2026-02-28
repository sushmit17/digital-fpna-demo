"""
Executor Agent 1 — Consolidation
Reads 5 unit files, consolidates into affiliate P&L, writes Excel output.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from tools.excel_parser import parse_unit_file
from tools.consolidator import consolidate

EXPECTED_UNITS = {
    "immunology":    "Immunology (Sales)",
    "oncology":      "Oncology (Sales)",
    "neurology":     "Neurology (Sales)",
    "rd":            "R&D",
    "general_admin": "General Admin",
}

def run(file_paths: list[str], output_path: str) -> dict:
    """
    file_paths: list of local paths to uploaded xlsx files
    Returns: {consolidated: dict, output_path: str, unit_names: list}
    """
    unit_data = []
    parsed_units = []
    for fp in file_paths:
        parsed = parse_unit_file(fp)
        unit_data.append(parsed)
        parsed_units.append(parsed["unit"])

    consolidated = consolidate(unit_data, output_path)
    return {
        "consolidated": consolidated,
        "output_path": output_path,
        "units_received": parsed_units,
    }
