"""
Executor Agent 2 — S&OP Deck
Generates CFO-quality variance narrative via Claude, then builds 5-slide PPT.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import anthropic
from tools.ppt_builder import build_deck

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

NARRATIVE_SYSTEM = """You are a senior FP&A Manager at a pharma commercial affiliate.
Write a concise, CFO-ready S&OP variance narrative based on the P&L data provided.
Tone: direct, factual, no filler. Use finance language.
Structure: 3–4 short paragraphs covering:
1. Overall performance summary (Net Sales, Division Margin vs Budget)
2. Sales performance by therapeutic area (Immunology, Oncology, Neurology)
3. Cost performance (R&D and SG&A vs Budget)
4. Key risks or opportunities going into next period
Keep total length to 180–220 words. Use EUR millions throughout."""

def run(consolidated_data: dict, output_path: str,
        unit_data: list = None) -> dict:
    """
    consolidated_data: output from executor_consolidation
    Returns: {narrative: str, ppt_path: str}
    """
    # Build a structured summary for the LLM
    lines_summary = []
    for line, vals in consolidated_data.items():
        if not isinstance(vals, dict):
                continue
        lbe = vals.get("LBE FY2026", 0)
        var_bgt = vals.get("Var vs Budget", 0)
        var_plbe = vals.get("Var vs Prior LBE", 0)
        var_py = vals.get("Var vs Prior Year", 0)
        lines_summary.append(
            f"  {line}: LBE={lbe:.1f}M | vs Budget={var_bgt:+.1f}M | "
            f"vs Prior LBE={var_plbe:+.1f}M | vs PY={var_py:+.1f}M"
        )

    prompt = f"""Commercial Affiliate P&L — LBE FY2026 (EUR millions):

{chr(10).join(lines_summary)}

Context:
- Immunology: strong performance, above Budget and Prior LBE
- Oncology: behind Budget due to delayed patient uptake in H1
- Neurology: ahead of Budget, growing vs Prior Year
- R&D: slight overrun vs Budget due to late-stage trial support
- SG&A: below Budget due to phasing of marketing spend

Write the S&OP narrative now."""

    msg = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=600,
        system=NARRATIVE_SYSTEM,
        messages=[{"role": "user", "content": prompt}]
    )
    narrative = msg.content[0].text.strip()

    # Build PPT
    ppt_path = build_deck(
        data=consolidated_data,
        narrative=narrative,
        output_path=output_path,
        unit_data=unit_data or [],
    )

    return {"narrative": narrative, "ppt_path": ppt_path}
