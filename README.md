# PBI AI Agent

> A locally-running AI agent for Power BI Desktop — reads your `.pbix` files directly, generates professional color themes, writes DAX measures, and provides intelligent dashboard design advice through a conversational web interface. No Azure. No cloud dependency. Completely free via Groq.

---

## The Key Idea

Power BI Desktop has no open API. Every third-party tool that claims to "work with Power BI" either requires an expensive Premium license, an Azure Active Directory setup, or only works on published cloud reports — never on local files.

This agent takes a different approach: it reads `.pbix` files directly as ZIP archives, parses the internal `Report/Layout` JSON to extract pages, visuals, field bindings, and dataset references, and feeds that structural context to a large language model. The result is an AI that genuinely knows your report — not a generic assistant that guesses.

The theme generation system works around Power BI's cryptographic file signing (introduced in Desktop v1.28) by producing standard JSON theme files that Power BI imports natively via its own trusted channel, rather than attempting direct file modification.

---

## Architecture & Flow

```
┌─────────────────────────────────────────────────────────────────────┐
│                        EXECUTION FLOW                               │
│                                                                     │
│  1. FILE INGESTION                                                  │
│     pbi_client.py opens the .pbix as a ZIP archive                 │
│     Reads Report/Layout (UTF-16-LE encoded JSON entry)             │
│     Extracts: pages · visual types · field bindings · titles       │
│     Resolves query aliases (d, d1...) to real table names          │
│     Falls back to visual-reference extraction when                  │
│     DataModelSchema is unavailable (Power BI v1.28+)               │
│                                                                     │
│  2. CONTEXT ASSEMBLY                                                │
│     app.py builds a structured snapshot of the report:             │
│     pages + visuals + measures + tables + relationships            │
│     This snapshot is injected into the LLM system prompt           │
│     so every chat response references your actual file             │
│                                                                     │
│  3. THEME GENERATION  (3 independent modes)                        │
│                                                                     │
│     [describe] → user inputs a natural language vibe string        │
│                  → LLM generates a full 18-key palette JSON        │
│                  → covers background, canvas, 6 chart colors,      │
│                     table header/rows, KPI cards, font, accent,    │
│                     positive/negative/neutral sentiment colors      │
│                                                                     │
│     [browse]   → page type detected from page name keywords        │
│                  → 5 curated category libraries matched:           │
│                     sales · marketing · finance · ops · executive  │
│                  → user selects from 4 hand-crafted palettes       │
│                                                                     │
│     [auto]     → LLM infers best palette from page name alone      │
│                  → no user input required                          │
│                                                                     │
│  4. THEME COMPILATION                                               │
│     theme_builder.py maps the raw palette dict                     │
│     to the strict Power BI theme JSON schema:                      │
│     name · dataColors · background · foreground · tableAccent      │
│     good · neutral · bad · maximum · center · minimum · null       │
│                                                                     │
│  5. DELIVERY                                                        │
│     Per-page ⬇ Download JSON button in the browser                │
│     ⬇ Download all as ZIP for bulk export                         │
│     Import in Desktop: View → Themes → Browse for themes          │
│                                                                     │
│  6. AI CHAT                                                         │
│     Persistent conversational session via Groq API                 │
│     Full report context injected on every request                  │
│     Supports: file attachments · images · CSV · JSON · TXT        │
│     JSON themes in replies get an auto ⬇ Download button          │
│     Conversation history maintained (last 20 turns)               │
└─────────────────────────────────────────────────────────────────────┘
```

---

## Module Responsibilities

| Module | Role |
|--------|------|
| `app.py` | Flask web server — routes, state management, API endpoints, download handlers |
| `pbi_client.py` | Deep `.pbix` reader — ZIP parsing, UTF-16-LE decoding, visual/field/measure extraction, alias resolution |
| `color_engine.py` | Theme generation — LLM prompting, palette parsing, curated library, color math (lighten, blend) |
| `theme_builder.py` | Palette compiler — maps raw palette dict to valid Power BI theme JSON schema |
| `agent_brain.py` | Conversational AI — persistent chat history, context injection, Groq API wrapper |
| `templates/index.html` | Dark web UI — tabbed interface, sidebar, chat panel, attachment upload, download buttons |
| `python_visuals/` | Premium visual templates — static dark dashboard scripts for Power BI Python visual |

---

## Web Interface

The frontend is a single-page dark application with six functional tabs:

| Tab | Function |
|-----|----------|
| **AI Agent** | Chat interface with file attachment support |
| **Load File** | Scan computer for `.pbix` files or paste a path |
| **Auto Theme** | One-click theme detection for all pages |
| **Describe** | Natural language theme generation with vibe chips |
| **Browse Palettes** | Category-matched curated palette selector |
| **Export / Apply** | Per-page download buttons and bulk ZIP export |

The sidebar displays the loaded file name, all report pages, and the full dataset schema — tables, columns, and detected measures — extracted live from the file.

---

## Prerequisites

- Python 3.9+
- Power BI Desktop (local installation)
- A free [Groq API key](https://console.groq.com) — no credit card required

---

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/pbi-ai-agent.git
cd pbi-ai-agent
pip install -r requirements.txt
```

Set your API key — either as an environment variable (recommended):

```bash
# Windows PowerShell
$env:GROQ_API_KEY = "gsk_..."

# Mac / Linux
export GROQ_API_KEY="gsk_..."
```

Or paste it directly in `app.py` line 16:

```python
GROQ_API_KEY = os.getenv("GROQ_API_KEY", "gsk_...")
```

---

## Usage

```bash
python app.py
```

Open **http://localhost:5000** in your browser.

### Recommended workflow

**Step 1 — Load your report**
Go to the **Load File** tab. Click *Scan computer for .pbix files* or paste the full path. The agent reads your entire report structure — pages, visuals, tables, measures — and displays them in the sidebar.

**Step 2 — Generate a theme**
Choose your mode:
- *Describe* — type `dark luxury navy gold` or `clean minimal white green`
- *Browse* — pick from curated palettes matched to your page type
- *Auto* — let the AI decide based on your page names

**Step 3 — Download and import**
Click **⬇ Download JSON** next to any generated theme, or **⬇ Download all as ZIP**. In Power BI Desktop: `View → Themes → Browse for themes → select the JSON file`.

**Step 4 — Chat with the agent**
Switch to the **AI Agent** tab. Ask anything:
- *"Review my DAX measures and tell me if they are correct"*
- *"What chart types should I use on Page 2?"*
- *"Write a measure for revenue month-over-month change"*
- *"Generate a dark royal blue theme for all pages"*

Attach screenshots, CSV files, or JSON directly in the chat input using the 📎 button.

---

## Python Visuals

For premium dark dashboard visuals inside Power BI Desktop:

```bash
pip install matplotlib numpy pandas
```

1. In Power BI Desktop: **Insert → Python visual**
2. Drag your data fields into the visual
3. Open `python_visuals/powerbi_python_visual.py`
4. Copy and paste the entire script into the Python script editor
5. Click **Run**

For AI-generated insights alongside the visual, use `powerbi_ai_visual.py` and set your Groq key on line 23.

---

## Why Not Direct `.pbix` Editing?

Power BI Desktop version 1.28+ (February 2026) cryptographically signs `.pbix` files using Windows DPAPI — the Data Protection API tied to your Windows user profile. Any external byte-level modification to the ZIP structure invalidates the signature and produces a `MashupValidationError` on open, corrupting the file.

The theme JSON import mechanism bypasses this entirely: Power BI validates and applies the theme file through its own trusted internal channel, making it the only reliable way to apply external styling changes to local reports without a Premium license.

---

## Project Structure

```
pbi-ai-agent/
├── app.py                          # Flask server — run this
├── agent_brain.py                  # Conversational AI logic
├── color_engine.py                 # Theme generation (3 modes + curated library)
├── theme_builder.py                # Palette → Power BI JSON compiler
├── pbi_client.py                   # Deep local .pbix reader
├── requirements.txt                # Dependencies
├── .env.example                    # Environment variable template
├── .gitignore
├── templates/
│   └── index.html                  # Dark web UI (single-page app)
└── python_visuals/
    ├── powerbi_python_visual.py    # Premium static dashboard visual
    └── powerbi_ai_visual.py        # AI-powered insights visual
```

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Backend | Python 3, Flask, Flask-CORS |
| AI / LLM | Groq API — `llama-3.3-70b-versatile` (free tier) |
| `.pbix` Parsing | Python `zipfile` stdlib + UTF-16-LE JSON decoding |
| Frontend | Vanilla HTML / CSS / JavaScript (no framework) |
| Python Visuals | matplotlib, numpy, pandas |
| Theme Delivery | Standard Power BI JSON theme import |

---

## License

MIT — free to use, modify, and distribute with attribution.
