"""
Validator Agent — checks outputs for completeness, controls and materiality.
Returns a structured validation report.
"""
import anthropic
import os

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

MATERIALITY_THRESHOLD = 5.0   # EUR millions — flag variances above this

PNL_REQUIRED_LINES = [
    "Net Sales", "Distribution Margin", "Total R&D",
    "Total SG&A", "Division Margin"
]

def validate_consolidation(consolidated_data: dict) -> dict:
    """
    Checks:
    1. All required P&L lines present
    2. Division Margin cross-check (DM - R&D - SGA = DivMar)
    3. Material variances flagged (>= MATERIALITY_THRESHOLD vs Budget)
    Returns: {passed: bool, flags: list[str], summary: str}
    """
    flags = []

    # Check 1: Required lines
    for line in PNL_REQUIRED_LINES:
        if line not in consolidated_data:
            flags.append(f"⚠️  MISSING LINE: '{line}' not found in consolidation")

    # Check 2: Division Margin cross-check
    dm  = consolidated_data.get("Distribution Margin", {}).get("LBE FY2026", 0)
    rd  = consolidated_data.get("Total R&D",            {}).get("LBE FY2026", 0)
    sga = consolidated_data.get("Total SG&A",           {}).get("LBE FY2026", 0)
    div = consolidated_data.get("Division Margin",      {}).get("LBE FY2026", 0)
    computed = round(dm - rd - sga, 1)
    if abs(computed - div) > 0.1:
        flags.append(f"❌  CROSS-CHECK FAIL: Division Margin {div} ≠ DM({dm}) - R&D({rd}) - SGA({sga}) = {computed}")
    else:
        flags.append(f"✅  Cross-check passed: Division Margin = {div} (DM {dm} − R&D {rd} − SGA {sga})")

    # Check 3: Material variances vs Budget
    for line, vals in consolidated_data.items():
        var_bgt = vals.get("Var vs Budget", 0)
        if abs(var_bgt) >= MATERIALITY_THRESHOLD:
            direction = "above" if var_bgt > 0 else "below"
            flags.append(f"🔴  MATERIAL: {line} is {abs(var_bgt):.1f}M {direction} Budget — requires explanation")

    passed = not any(f.startswith("❌") or f.startswith("⚠️") for f in flags)

    summary = (
        f"Validation {'PASSED ✅' if passed else 'FAILED ❌'} — "
        f"{len([f for f in flags if f.startswith('🔴')])} material variances flagged, "
        f"{len([f for f in flags if f.startswith('❌')])} errors found."
    )
    return {"passed": passed, "flags": flags, "summary": summary}


def validate_sop_deck(narrative: str, consolidated_data: dict) -> dict:
    """
    Uses Claude to check narrative quality:
    - Every material variance commented?
    - Numbers consistent with source data?
    - Tone appropriate for CFO audience?
    """
    material_lines = [
        f"{line}: Var vs Budget = {vals.get('Var vs Budget', 0):.1f}M"
        for line, vals in consolidated_data.items()
        if abs(vals.get("Var vs Budget", 0)) >= MATERIALITY_THRESHOLD
    ]

    prompt = f"""You are a finance controller reviewing an S&OP narrative for a CFO.

Material variances that MUST be explained:
{chr(10).join(material_lines)}

Narrative to review:
{narrative}

Check:
1. Does the narrative address each material variance above?
2. Are there any numerical inconsistencies?
3. Is the tone suitable for a CFO audience?

Respond in this exact format:
RESULT: PASS or FAIL
ISSUES: [bullet list of issues, or "None"]
"""
    msg = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=400,
        messages=[{"role": "user", "content": prompt}]
    )
    response = msg.content[0].text.strip()
    passed = response.startswith("RESULT: PASS")
    return {
        "passed": passed,
        "review": response,
        "summary": f"Narrative validation {'PASSED ✅' if passed else 'NEEDS REVIEW ⚠️'}"
    }
