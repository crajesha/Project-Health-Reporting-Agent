# RAG Methodology — Project Health Reporting Agent

## Principle

RAG status is **calculated independently from raw task data**, not copied from
the plan's existing `Schedule Health` / `RAG` column. The existing field is
treated as one more input to compare against, not the source of truth —
different PMs apply it inconsistently, and it reflects task-level judgment
rather than a portfolio-comparable score. Every status this agent produces
comes with the specific numbers that drove it, so a VP or PM can trace *why*,
not just *what*.

## The five signals

| # | Signal | Weight | What it measures | Data used |
|---|--------|--------|-------------------|-----------|
| 1 | **Schedule Slippage** | 30% | How far behind plan tasks are | `Variance` (days late), tasks with `End Date` in the past that aren't `Completed` |
| 2 | **Completion** | 20% | Real progress vs. time elapsed | `% Complete`, plus a penalty if completion trails the % of the timeline already elapsed |
| 3 | **Critical-Task Health** | 20% | Risk concentrated on the tasks that most threaten the end date | `Critical?` flag, cross-referenced with each critical task's own `Schedule Health` and `Variance` |
| 4 | **Blockers / Sentiment** | 15% | Leading indicators that precede slippage | Free-text `Comments` sheet and `Status Comment` column, scanned for blocker/risk/pending language |
| 5 | **Milestone Health** | 15% | Phase-level delivery, not just task counts | `Phase/Milestone` rows: % completed and % showing variance |

Each signal is scored 0 (worst) – 100 (best) using explicit, documented rules
(see `src/rag_engine.py`), then combined into a **weighted composite score**.

## Score → RAG mapping

| Composite Score | Status |
|---|---|
| ≥ 80 | 🟢 Green |
| 60 – 79 | 🟡 Amber |
| < 60 | 🔴 Red |

## Override rules (guardrails)

A weighted average alone can hide concentrated risk — e.g. 90% of tasks could
be healthy while the 10% that are critical-path are on fire. Two override
rules cap the rating regardless of the composite:

- **≥ 30% of critical tasks internally rated Red** → status capped at **Red**.
- **10–29% of critical tasks rated Red** → status capped at **Amber**.
- **≥ 60% of logged comments flag blockers/pending items** (min. 3 comments)
  → status capped at **Amber**.

## Comparing to the sheet's self-reported status

For every project, the agent also reports whether its computed status
**agrees or disagrees** with the workbook's own `Schedule Health` value, and
explains the disagreement in plain English (e.g. *"individual tasks look
fine, but 25 open tasks are already past their planned end date and 2 of 50
critical-path tasks are Red — that combination pulls the overall picture to
Amber even though the sheet says Red / Green"*). This is deliberate: the
assignment brief asks the agent not to simply mirror the existing field.

## Confidence score

A 0–1 confidence score reflects **how much of the underlying data was
actually available** (variance data, completion %, critical flags, comments,
milestones). A project with sparse comments and no milestone tracking still
gets a status, but with a lower confidence score — flagging to the reader
that the rating rests on fewer signals than usual.

## Assumptions made about the data

- `Variance` values follow the `-Nd` / `Nd` text convention observed in the
  source sheets and are parsed to signed integer days.
- `#UNPARSEABLE`, blank cells, and `NaN` are all treated as **missing data**,
  not as zero or false — a missing `Critical?` flag means "not critical," not
  "critical = 0 risk."
- `% Complete` is normalized to a 0–1 scale regardless of whether the source
  stored it as a fraction or a whole percentage.
- The task sheet, `Comments` sheet, and `Summary` sheet are identified by
  role (widest non-reserved sheet / exact name match), not by a fixed sheet
  index, since sheet names vary per project (e.g. "Project Plan" vs.
  "Outokumpu- S2P Project").
- Where a project has no comments or no milestone rows at all, that signal is
  scored neutrally (not automatically penalized) but flagged in the
  confidence score, since absence of data is not evidence of good health.
- Weights (30/20/20/15/15) and thresholds (80/60) are configurable in
  `RagEngine.__init__` — they encode a Professional-Services-reasonable
  starting point, not a fixed law, and should be tuned against a larger
  historical sample as more projects run through the agent.
