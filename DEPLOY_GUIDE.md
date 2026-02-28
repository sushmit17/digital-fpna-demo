# Digital FP&A Manager — Demo
## Deploy Guide: GitHub → Render.com in 10 minutes

---

### What you'll need
- A free **GitHub** account — github.com
- A free **Render.com** account — render.com
- An **Anthropic API key** — console.anthropic.com (takes 2 minutes to create)

---

## Step 1 — Put the code on GitHub

1. Go to **github.com** → click **New repository**
2. Name it `digital-fpna-demo` → click **Create repository**
3. Upload all the project files (drag & drop the folder contents into GitHub)
4. Click **Commit changes**

Your repo is now live at `github.com/YOUR_USERNAME/digital-fpna-demo`

---

## Step 2 — Deploy on Render

1. Go to **render.com** → Sign up (free) → click **New +** → **Web Service**
2. Connect your GitHub account and select `digital-fpna-demo`
3. Fill in the settings:

| Setting | Value |
|---|---|
| **Name** | `digital-fpna-demo` |
| **Runtime** | `Python 3` |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `chainlit run app.py --host 0.0.0.0 --port $PORT` |
| **Instance Type** | Free |

4. Click **Advanced** → **Add Environment Variable**:
   - Key: `ANTHROPIC_API_KEY`
   - Value: `sk-ant-...` *(paste your key)*

5. Click **Create Web Service**

Render will build and deploy — takes about 2 minutes.

---

## Step 3 — Share the link

Once deployed, Render gives you a permanent URL like:

```
https://digital-fpna-demo.onrender.com
```

Share this link with your audience. Anyone can open it in a browser — **no login, no install required**.

---

## Step 4 — Running the demos

**Prepare the test files**

The `input_files/` folder in your repo contains 5 pre-built Excel files:
- `immunology_LBE_FY2026.xlsx`
- `oncology_LBE_FY2026.xlsx`
- `neurology_LBE_FY2026.xlsx`
- `rd_LBE_FY2026.xlsx`
- `general_admin_LBE_FY2026.xlsx`

Download these and keep them ready for the demo.

---

**Demo 1 — LBE Consolidation (~30 seconds)**

1. Open the link
2. Type: `run consolidation`
3. Upload all 5 Excel files when prompted
4. Watch the agent consolidate, cross-check, and flag variances
5. Download the consolidated P&L Excel file

**Demo 2 — S&OP Deck (~20 seconds)**

1. After Demo 1 completes, type: `generate S&OP deck`
2. Watch the agent write the narrative and build the presentation
3. Download the 5-slide PowerPoint

---

## Suggested demo script for finance leadership

> *"I'm going to show you two things the Digital FP&A Manager agent can do. First, I'll run the LBE consolidation — something that normally takes half a day. Then I'll ask it to generate the S&OP deck."*

**[Type 'run consolidation' and upload the 5 files]**

> *"Notice the Planner Agent reading the intent and routing to the right executor. The agent knows which files to expect for this cycle. Now watch the Validator — it's cross-checking Division Margin and flagging the Oncology variance automatically."*

**[After Demo 1 completes]**

> *"Now let's go straight to the S&OP deck."*

**[Type 'generate S&OP deck']**

> *"The agent is now writing the CFO narrative and building all five slides from the data it just consolidated. No copy-paste, no manual formatting."*

**[Download and open the PPT]**

> *"This is the output we could use in the S&OP meeting. The architecture, the controls, the governance — it all mirrors what we've designed in the full system. The difference in production is Microsoft Teams instead of this chat interface, and your actual ERP data instead of these test files."*

---

## Cost

The demo costs approximately **EUR 0.002 per full run** (Demo 1 + Demo 2) using Claude Haiku.
Running it 100 times costs about EUR 0.20.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Site shows "Service unavailable" | Free tier spins down after 15 min. Wait 30 seconds and refresh. |
| "API key invalid" error | Check the env var in Render dashboard — no spaces around the key |
| File upload not working | Ensure files are `.xlsx` format, not `.xls` or `.csv` |
| Demo 2 says "no data found" | Run Demo 1 first in the same session |

---

*Built with: Chainlit · LangGraph · Anthropic Claude · python-pptx · openpyxl*
*Hosted on: Render.com (free tier)*
