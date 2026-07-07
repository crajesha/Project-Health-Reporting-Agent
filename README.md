# Project Health Reporting Agent

An AI-assisted system that reads raw project-plan workbooks, computes an
independent RAG (Red/Amber/Green) status for each project, explains the
status in plain English, drafts weekly reports, generates a monthly
executive presentation that finds *trends* across the portfolio, and
suggests (never sends) follow-up emails for delayed tasks.

Built against the two sample workbooks provided (`Project_Plan_B.xlsx`,
`S2P_Project.xlsx`) and designed to generalize to any workbook that follows
the same task-sheet / Comments / Summary convention.

---

## 1. What it does

| Requirement | Delivered as |
|---|---|
| Read project plans from Excel | `src/excel_parser.py` |
| Determine RAG status (not copied from the sheet) | `src/rag_engine.py` |
| Plain-English reasoning | `src/reasoning.py` |
| Handle messy/incomplete data | built into the parser + engine (see below) |
| Weekly reports | `src/report_generator.py` → `outputs/weekly/*.md` / `*.html` |
| Weekly scheduler (bonus) | `src/scheduler.py` (APScheduler, cron: every Monday 07:00) |
| Monthly executive presentation (5–7 slides, trend-focused) | `src/build_pptx.js` → `outputs/presentation/*.pptx` |
| Executive Action Assistant (bonus) | `src/action_assistant.py` — drafts follow-up emails, never sends |
| Dashboard | `src/dashboard.py` (Streamlit) |

---

## 2. Architecture

```
Excel workbooks (.xlsx)
        │
        ▼
┌─────────────────┐     dataclasses: ParsedProject, ProjectSummary, CommentEntry
│  excel_parser.py │ ──▶ (task sheet, Comments sheet, Summary sheet;
└─────────────────┘      cleans #UNPARSEABLE / NaN / blank / date parsing)
        │
        ▼
┌─────────────────┐     5 weighted, auditable signals → composite score →
│  rag_engine.py   │ ──▶ RAG + override guardrails + confidence + comparison
└─────────────────┘      against the sheet's own self-reported status
        │
        ▼
┌─────────────────┐     Executive summary / risk explanation / positives /
│  reasoning.py    │ ──▶ concerns / recommendations, in a senior-delivery-
└─────────────────┘      manager voice (template-driven NLG, deterministic)
        │
   ┌────┴─────┬──────────────────────┐
   ▼           ▼                     ▼
report_generator.py   action_assistant.py     (portfolio_data.json)
   │ weekly .md/.html    │ suggested email          │
   │                     │ drafts (not sent)         ▼
   ▼                     ▼                  build_pptx.js
outputs/weekly/     outputs/actions/        outputs/presentation/*.pptx
                                             (5-7 slide executive deck)
```

`main.py` orchestrates the whole pipeline for one or many input files;
`scheduler.py` wraps it in a weekly cron job; `dashboard.py` wraps it in an
interactive Streamlit UI.

---

## 3. Folder structure

```
project_health_agent/
├── README.md
├── requirements.txt
├── data/                          # sample input workbooks
│   ├── Project_Plan_B.xlsx
│   └── S2P_Project.xlsx
├── docs/
│   └── RAG_Methodology.md         # Phase 1 deliverable (one-pager)
├── src/
│   ├── excel_parser.py            # Step 2
│   ├── rag_engine.py              # Step 3 — deterministic scoring
│   ├── reasoning.py                # Step 4 — plain-English narrative
│   ├── report_generator.py        # Step 5 — weekly md/html reports
│   ├── action_assistant.py        # bonus — Executive Action Assistant
│   ├── build_pptx.js              # Step 7 — executive presentation
│   ├── dashboard.py                # Step 6 — Streamlit dashboard
│   ├── scheduler.py                # Step 8 — weekly APScheduler job
│   └── main.py                     # CLI orchestrator
└── outputs/
    ├── weekly/                    # generated per-project reports
    ├── actions/                   # generated follow-up email drafts
    ├── presentation/              # generated executive .pptx
    └── portfolio_data.json        # intermediate aggregate data
```

---

## 4. Installation

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# The executive deck is generated with pptxgenjs (Node)
npm install -g pptxgenjs
```

## 5. How to run

**Run the full pipeline once** (parses every workbook, writes weekly
reports, drafted follow-ups, and rebuilds the executive deck):

```bash
python src/main.py --input "data/*.xlsx" --out outputs
```

**Run the interactive dashboard:**

```bash
streamlit run src/dashboard.py
```

**Run the weekly scheduler** (long-lived process; runs an immediate pass on
startup, then every Monday at 07:00):

```bash
python src/scheduler.py
```

Environment variables for the scheduler:
- `PHA_INPUT_GLOB` (default `data/*.xlsx`)
- `PHA_OUTPUT_DIR` (default `outputs`)

---

## 6. Sample outputs (generated from the two provided workbooks)

| Project | Computed RAG | Sheet's own Schedule Health | Agree? |
|---|---|---|---|
| Project_Plan_B | 🟡 Amber (76.5/100) | Red | No — see explanation in report |
| S2P Project (Outokumpu) | 🔴 Red (58.7/100) | Green | No — see explanation in report |

Both disagreements are intentional and explained in each weekly report:
neither project's status was copied from the sheet — see
`outputs/weekly/*.md` and `docs/RAG_Methodology.md`.

---

## 7. Design decisions

**Why deterministic scoring instead of an LLM for the RAG calculation?**
RAG status drives governance and client conversations — it needs to be
reproducible, explainable, and auditable line-by-line. `rag_engine.py` is
pure Python with no external API dependency, so the same input always
produces the same output, and every score can be traced back to the exact
numbers that produced it (see the `facts` dict on every `MetricResult`).

**Why isn't the plain-English narrative generated by a hosted LLM (Gemini /
Claude)?** Two reasons. First, it keeps the agent runnable with zero API
keys and zero external dependency risk — a real constraint for a
Professional-Services tool that needs to work reliably on a weekly cadence.
Second, template-driven generation over the *same* structured facts the RAG
engine already computed means the narrative can never contradict the score
— it's built from the same numbers, not a separate model's guess at them.
`reasoning.py` is written as a clean interface (`build_narrative(project,
result) -> Narrative`) specifically so a hosted LLM call can be dropped in
behind it later for richer prose, without touching the scoring engine.

**Why not simply copy the sheet's `Schedule Health` / `RAG` column?** The
assignment explicitly calls this out, and the sample data shows why it
matters: `Project_Plan_B`'s sheet says **Red** while the computed status is
**Amber**, and `S2P_Project`'s sheet says **Green** while the computed
status is **Red** — the opposite direction. Column-level status reflects
one PM's local judgment on one day; the composite score is comparable across
PMs and projects.

**Why "suggest emails," never "send emails"?** Executive communication
should stay under human control. The Action Assistant produces a
subject/body draft with the owner and reason pre-filled; a VP reviews,
edits, and sends it from their own mail client. Nothing in this codebase
sends email.

**Why python-pptx skill's pptxgenjs pipeline instead of a template?** No
existing branded template was provided, so the deck is built from scratch
following Anthropic's internal design skill: one dominant color (Midnight
Executive navy), no decorative stripes, varied layouts per slide (cards,
doughnut + bar charts, numbered trend list, risk cards, table), 16:9 wide
format suitable for a client-facing screen share.

---

## 8. Assumptions about the data

See `docs/RAG_Methodology.md` §"Assumptions made about the data" for the
full list. Highlights:
- `#UNPARSEABLE`, blank cells, and `NaN` are all treated as **missing**, not
  as zero/false.
- The task-level sheet is auto-detected by process of elimination (whatever
  isn't named `Comments` or `Summary`), since its name varies per project.
- `Variance` values follow the observed `-Nd` / `Nd` text convention.
- Absence of comments or milestone data lowers the *confidence* score rather
  than silently penalizing the RAG score itself.

## 9. Known limitations / future enhancements

- The narrative generator is template-driven; swapping in a hosted LLM
  behind `reasoning.build_narrative()` would allow more varied prose while
  keeping the same structured inputs.
- With only two sample projects, "trend" analysis on the executive deck is
  necessarily thinner than it will be once a real multi-month, multi-project
  history accumulates — the deck is built to scale to that history without
  code changes (it operates on `portfolio_data.json`, not hardcoded values).
- PDF export of the weekly report is HTML with print-friendly CSS today;
  wiring in a headless-Chrome step (`playwright` / `weasyprint`) would give
  a one-command PDF.
- The Executive Action Assistant currently drafts one email per overdue
  task; a natural next step is grouping multiple overdue tasks for the same
  owner into a single digest email.
- Owner/email extraction relies on the `Assigned To` / `Owner` columns
  containing an email address or a recognizable name; workbooks that only
  list role names (e.g. "Zycus Project Team") will get a role-addressed
  draft rather than a named recipient.
