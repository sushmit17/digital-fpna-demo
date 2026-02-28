"""
Planner Agent — interprets user intent and routes to Demo 1 or Demo 2.
Uses Claude to classify intent then returns a routing decision.
"""
import anthropic
import os

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

SYSTEM = """You are the Planner Agent for the Digital FP&A Manager system.
Your job is to read the user's message and decide which workflow to run.

Respond with ONLY one of these exact strings — nothing else:
- DEMO1_CONSOLIDATION   (user wants to consolidate LBE files / run consolidation)
- DEMO2_SOP             (user wants to generate S&OP deck / presentation / slides)
- CLARIFY               (intent is unclear — ask for clarification)
"""

def plan(user_message: str) -> str:
    """Returns: 'DEMO1_CONSOLIDATION', 'DEMO2_SOP', or 'CLARIFY'"""
    msg = client.messages.create(
        model="claude-haiku-4-5",
        max_tokens=20,
        system=SYSTEM,
        messages=[{"role": "user", "content": user_message}]
    )
    result = msg.content[0].text.strip()
    if result not in ("DEMO1_CONSOLIDATION", "DEMO2_SOP", "CLARIFY"):
        return "CLARIFY"
    return result
