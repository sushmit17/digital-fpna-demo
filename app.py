"""
Digital FP&A Manager — Demo
Chainlit chat application orchestrating Planner → Executor → Validator agents.
"""
import os
import sys
import tempfile
import asyncio
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

import chainlit as cl

sys.path.insert(0, os.path.dirname(__file__))
from agents.planner import plan
from agents.validator import validate_consolidation, validate_sop_deck
from agents import executor_consolidation, executor_sop

# ── State (per session) ───────────────────────────────────────────────────────
WELCOME = """\
# 🏥 Digital FP&A Manager — Demo

Welcome. I am the **Digital FP&A Manager Agent** for your Commercial Affiliate.

I can run two demos today:

---

**Demo 1 — LBE Consolidation**
I will consolidate submissions from your 5 business units into a single \
Affiliate P&L with variances vs Prior LBE, Budget and Prior Year.

👉 Type **"run consolidation"** and upload all 5 unit files when prompted.

---

**Demo 2 — S&OP Deck Generation**
I will generate a CFO-ready variance narrative and a 5-slide PowerPoint \
S&OP presentation from the consolidated data.

👉 Run Demo 1 first, then type **"generate S&OP deck"**.

---

📁 **Test files are provided** in the `input_files/` folder of this project.
"""

DEMO1_PROMPT = """\
📂 **Upload your 5 unit submission files now.**

I am expecting:
| # | Unit | File |
|---|------|------|
| 1 | Immunology (Sales) | `immunology_LBE_FY2026.xlsx` |
| 2 | Oncology (Sales) | `oncology_LBE_FY2026.xlsx` |
| 3 | Neurology (Sales) | `neurology_LBE_FY2026.xlsx` |
| 4 | R&D | `rd_LBE_FY2026.xlsx` |
| 5 | General Admin | `general_admin_LBE_FY2026.xlsx` |

Upload all 5 files and I will begin the consolidation.
"""


@cl.on_chat_start
async def on_start():
    cl.user_session.set("consolidated_data", None)
    cl.user_session.set("unit_sales", None)
    cl.user_session.set("awaiting_files", False)
    await cl.Message(content=WELCOME).send()


@cl.on_message
async def on_message(message: cl.Message):
    awaiting = cl.user_session.get("awaiting_files", False)

    # ── File submission path ──────────────────────────────────────────────────
    if awaiting and message.elements:
        await handle_file_upload(message)
        return

    # ── Intent routing ────────────────────────────────────────────────────────
    thinking = cl.Message(content="")
    await thinking.send()

    step_msg = await cl.Message(
        content="🧠 **Planner Agent** — reading your intent…"
    ).send()

    route = plan(message.content)

    if route == "DEMO1_CONSOLIDATION":
        await step_msg.update()
        await cl.Message(
            content=(
                "🧠 **Planner Agent** → routing to **Consolidation Executor**\n\n"
                + DEMO1_PROMPT
            )
        ).send()
        cl.user_session.set("awaiting_files", True)

    elif route == "DEMO2_SOP":
        consolidated = cl.user_session.get("consolidated_data")
        if not consolidated:
            await cl.Message(
                content=(
                    "⚠️ No consolidated data found yet.\n\n"
                    "Please run **Demo 1** first by typing *'run consolidation'* "
                    "and uploading the 5 unit files."
                )
            ).send()
            return
        await step_msg.update()
        await run_demo2(consolidated)

    else:
        await cl.Message(
            content=(
                "I didn't quite catch that. Try:\n"
                "- **'run consolidation'** — to start Demo 1\n"
                "- **'generate S&OP deck'** — to start Demo 2 (after Demo 1)"
            )
        ).send()


async def handle_file_upload(message: cl.Message):
    """Process uploaded unit files and run Demo 1."""
    files = message.elements
    cl.user_session.set("awaiting_files", False)

    await cl.Message(
        content=f"📥 Received **{len(files)} file(s)**. Starting consolidation…"
    ).send()

    # Save uploaded files to temp dir
    tmp_dir = tempfile.mkdtemp()
    file_paths = []
    for f in files:
        dest = os.path.join(tmp_dir, Path(f.name).name)
        # Chainlit stores file at f.path
        import shutil
        shutil.copy(f.path, dest)
        file_paths.append(dest)

    await run_demo1(file_paths, tmp_dir)


async def run_demo1(file_paths: list, tmp_dir: str):
    """Execute the consolidation workflow."""

    # ── Step 1: Parse & Consolidate ───────────────────────────────────────────
    step1 = await cl.Message(
        content="⚙️ **Consolidation Executor** — parsing unit files…"
    ).send()

    loop = asyncio.get_event_loop()

    try:
        output_xlsx = os.path.join(tmp_dir, "Affiliate_Consolidated_LBE_FY2026.xlsx")
        result = await loop.run_in_executor(
            None, executor_consolidation.run, file_paths, output_xlsx
        )
    except Exception as e:
        await cl.Message(content=f"❌ Consolidation error: {e}").send()
        return

    consolidated = result["consolidated"]
    units = result["units_received"]

    await cl.Message(
        content=(
            f"⚙️ **Consolidation Executor** — files parsed ✅\n\n"
            f"Units received: {', '.join(units)}\n\n"
            f"Building consolidated P&L…"
        )
    ).send()

    # ── Step 2: Validator ─────────────────────────────────────────────────────
    await cl.Message(
        content="🔍 **Validator Agent** — running controls & cross-checks…"
    ).send()

    validation = validate_consolidation(consolidated)

    flags_text = "\n".join(f"  {f}" for f in validation["flags"])
    await cl.Message(
        content=(
            f"🔍 **Validator Agent** — {validation['summary']}\n\n"
            f"```\n{flags_text}\n```"
        )
    ).send()

    # ── Step 3: Output ────────────────────────────────────────────────────────
    # Store for Demo 2
    cl.user_session.set("consolidated_data", consolidated)

    # Build a readable P&L summary for chat
    pnl_lines = [
        "| P&L Line | LBE | vs Prior LBE | vs Budget | vs PY |",
        "|---|---|---|---|---|",
    ]
    order = ["Net Sales", "Distribution Margin", "Total R&D",
             "Total SG&A", "Division Margin"]
    for line in order:
        if line in consolidated:
            v = consolidated[line]
            def fmt(x, sign=False):
                if x == 0:
                    return "-"
                s = f"+{x:.1f}" if (sign and x > 0) else f"{x:.1f}"
                return s
            pnl_lines.append(
                f"| **{line}** | {fmt(v['LBE FY2026'])} "
                f"| {fmt(v['Var vs Prior LBE'], True)} "
                f"| {fmt(v['Var vs Budget'], True)} "
                f"| {fmt(v['Var vs Prior Year'], True)} |"
            )

    await cl.Message(
        content=(
            "✅ **Consolidation complete!**\n\n"
            + "\n".join(pnl_lines)
            + "\n\n*All figures EUR millions*"
        )
    ).send()

    # Send Excel file
    await cl.Message(
        content="📊 **Consolidated P&L file ready for download:**",
        elements=[
            cl.File(
                name="Affiliate_Consolidated_LBE_FY2026.xlsx",
                path=output_xlsx,
                display="inline",
            )
        ]
    ).send()

    await cl.Message(
        content=(
            "---\n"
            "Demo 1 complete. You can now type **'generate S&OP deck'** "
            "to run Demo 2 and get the PowerPoint presentation."
        )
    ).send()


async def run_demo2(consolidated: dict):
    """Execute the S&OP deck generation workflow."""

    # ── Step 1: Narrative generation ─────────────────────────────────────────
    await cl.Message(
        content=(
            "🧠 **Planner Agent** → routing to **S&OP Executor**\n\n"
            "⚙️ **S&OP Executor** — generating variance narrative…"
        )
    ).send()

    tmp_dir = tempfile.mkdtemp()
    ppt_path = os.path.join(tmp_dir, "Affiliate_SOP_Deck_FY2026.pptx")

    loop = asyncio.get_event_loop()

    try:
        result = await loop.run_in_executor(
            None, executor_sop.run, consolidated, ppt_path, None
        )
    except Exception as e:
        await cl.Message(content=f"❌ S&OP generation error: {e}").send()
        return

    narrative = result["narrative"]

    # ── Step 2: Validator ─────────────────────────────────────────────────────
    await cl.Message(
        content="🔍 **Validator Agent** — checking narrative coverage & consistency…"
    ).send()

    val = await loop.run_in_executor(
        None, validate_sop_deck, narrative, consolidated
    )

    await cl.Message(
        content=f"🔍 **Validator Agent** — {val['summary']}"
    ).send()

    # ── Step 3: Output ────────────────────────────────────────────────────────
    await cl.Message(
        content=(
            "✅ **S&OP narrative generated:**\n\n"
            "---\n\n"
            + narrative
            + "\n\n---"
        )
    ).send()

    await cl.Message(
        content="📑 **S&OP PowerPoint deck ready for download:**",
        elements=[
            cl.File(
                name="Affiliate_SOP_Deck_FY2026.pptx",
                path=ppt_path,
                display="inline",
            )
        ]
    ).send()

    await cl.Message(
        content=(
            "**Deck contains 5 slides:**\n"
            "- Slide 1: Executive Summary & KPIs\n"
            "- Slide 2: Sales Forecast by Therapeutic Area\n"
            "- Slide 3: P&L vs Prior LBE\n"
            "- Slide 4: P&L vs Budget\n"
            "- Slide 5: P&L vs Prior Year\n\n"
            "*Demo complete. You may restart the conversation to run again.*"
        )
    ).send()
