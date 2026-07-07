"""
main.py
-------
CLI entry point for the Project Health Reporting Agent.

Usage:
    python src/main.py --input data/*.xlsx --out outputs

Runs the full pipeline for every input workbook:
    parse -> score RAG -> generate reasoning -> write weekly report
    -> draft follow-up action items -> (optionally) build the
       monthly executive presentation from the whole portfolio.
"""

from __future__ import annotations

import argparse
import glob
import json
import logging
import os
import subprocess
import sys

from excel_parser import ExcelParser
from rag_engine import RagEngine
from reasoning import build_narrative
from report_generator import save_weekly_report
from action_assistant import build_action_items, render_action_items_markdown

logger = logging.getLogger("main")


def run_pipeline(input_files: list[str], out_dir: str, build_deck: bool = True) -> list[dict]:
    parser = ExcelParser()
    engine = RagEngine()

    weekly_dir = os.path.join(out_dir, "weekly")
    actions_dir = os.path.join(out_dir, "actions")
    os.makedirs(weekly_dir, exist_ok=True)
    os.makedirs(actions_dir, exist_ok=True)

    portfolio = []
    for filepath in input_files:
        try:
            project = parser.parse_file(filepath)
        except Exception as exc:
            logger.error("Skipping %s: failed to parse (%s)", filepath, exc)
            continue

        result = engine.evaluate(project)
        narrative = build_narrative(project, result)
        paths = save_weekly_report(project, result, narrative, weekly_dir)

        action_items = build_action_items(project.project_name, project.tasks)
        actions_md = render_action_items_markdown(project.project_name, action_items)
        safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in project.project_name).strip().replace(" ", "_")
        actions_path = os.path.join(actions_dir, f"{safe_name}_suggested_followups.md")
        with open(actions_path, "w") as f:
            f.write(actions_md)

        m = {mm.name: mm for mm in result.metrics}
        portfolio.append({
            "name": project.project_name,
            "pm": project.summary.project_manager,
            "rag": result.rag,
            "score": result.composite_score,
            "confidence": result.confidence,
            "sheet_status": result.sheet_reported_status,
            "agrees": result.status_agrees_with_sheet,
            "pct_complete": m["completion"].facts["avg_pct_complete"],
            "pct_delayed_tasks": m["schedule_slippage"].facts["pct_delayed_tasks"],
            "missed_end_dates": m["schedule_slippage"].facts["missed_end_dates"],
            "avg_delay_days": m["schedule_slippage"].facts["avg_delay_days"],
            "critical_count": m["critical_tasks"].facts.get("critical_count", 0),
            "critical_red": m["critical_tasks"].facts.get("critical_red", 0),
            "critical_delayed": m["critical_tasks"].facts.get("critical_delayed", 0),
            "flagged_comments": m["blockers"].facts.get("flagged_count", 0),
            "total_comments": m["blockers"].facts.get("comment_count", 0),
            "milestones_completed": m["milestones"].facts.get("milestones_completed", 0),
            "milestones_total": m["milestones"].facts.get("milestone_count", 0),
            "total_tasks": len(project.tasks),
            "stage": project.summary.project_stage,
            "weekly_report_md": paths["markdown"],
            "weekly_report_html": paths["html"],
            "suggested_followups": actions_path,
        })
        logger.info("Processed %s -> %s (%.1f/100)", project.project_name, result.rag, result.composite_score)

    portfolio_json_path = os.path.join(out_dir, "portfolio_data.json")
    with open(portfolio_json_path, "w") as f:
        json.dump(portfolio, f, indent=2)

    if build_deck and portfolio:
        deck_dir = os.path.join(out_dir, "presentation")
        os.makedirs(deck_dir, exist_ok=True)
        deck_path = os.path.join(deck_dir, "Executive_Project_Health_Review.pptx")
        script_path = os.path.join(os.path.dirname(__file__), "build_pptx.js")
        try:
            env = os.environ.copy()
            env["NODE_PATH"] = subprocess.check_output(["npm", "root", "-g"], text=True).strip()
            subprocess.run(["node", script_path, portfolio_json_path, deck_path], check=True, env=env)
            logger.info("Executive presentation written to %s", deck_path)
        except Exception as exc:
            logger.error("Presentation generation failed: %s", exc)

    return portfolio


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
    ap = argparse.ArgumentParser(description="Project Health Reporting Agent")
    ap.add_argument("--input", nargs="+", required=True, help="Glob(s) or explicit paths to .xlsx project plans")
    ap.add_argument("--out", default="outputs", help="Output directory")
    ap.add_argument("--no-deck", action="store_true", help="Skip building the executive PowerPoint")
    args = ap.parse_args()

    files = []
    for pattern in args.input:
        matches = glob.glob(pattern)
        files.extend(matches if matches else [pattern])
    files = sorted(set(files))

    if not files:
        logger.error("No input files matched: %s", args.input)
        sys.exit(1)

    portfolio = run_pipeline(files, args.out, build_deck=not args.no_deck)
    print(json.dumps([{"project": p["name"], "rag": p["rag"], "score": p["score"]} for p in portfolio], indent=2))


if __name__ == "__main__":
    main()
