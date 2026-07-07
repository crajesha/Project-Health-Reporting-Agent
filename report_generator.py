"""
report_generator.py
--------------------
Produces the weekly per-project health report in Markdown and HTML
(PDF-ready: the HTML uses print-friendly CSS and can be converted to
PDF with any headless-Chrome/wkhtmltopdf step in a CI pipeline).
"""

from __future__ import annotations

import os
from datetime import datetime

from excel_parser import ParsedProject
from rag_engine import RagResult
from reasoning import Narrative

RAG_COLOR = {"Green": "#1a7f37", "Amber": "#b98400", "Red": "#cf222e"}
RAG_EMOJI = {"Green": "🟢", "Amber": "🟡", "Red": "🔴"}


def _fmt_date(ts):
    return ts.strftime("%d %b %Y") if ts is not None else "N/A"


def render_markdown(project: ParsedProject, result: RagResult, narrative: Narrative) -> str:
    m = {mm.name: mm for mm in result.metrics}
    lines = []
    lines.append(f"# Weekly Project Health Report — {project.project_name}")
    lines.append("")
    lines.append(f"**Report date:** {_fmt_date(result.as_of)}  ")
    lines.append(f"**Project Manager:** {project.summary.project_manager or 'N/A'}  ")
    lines.append(f"**Overall Status:** {RAG_EMOJI[result.rag]} **{result.rag}**  "
                  f"(composite score {result.composite_score}/100, confidence {int(result.confidence * 100)}%)")
    lines.append("")

    if not result.status_agrees_with_sheet:
        lines.append(f"> ⚠️ **Independent status differs from source sheet.** The workbook's Schedule "
                      f"Health field reads **{result.sheet_reported_status}**; this agent calculates "
                      f"**{result.rag}** based on a multi-signal composite. See Risk Explanation below.")
        lines.append("")

    lines.append("## Executive Summary")
    lines.append(narrative.executive_summary)
    lines.append("")

    lines.append("## Risk Explanation")
    lines.append(narrative.risk_explanation)
    lines.append("")

    lines.append("## Positive Observations")
    for p in narrative.positive_observations:
        lines.append(f"- {p}")
    lines.append("")

    lines.append("## Top Concerns")
    for c in narrative.top_concerns:
        lines.append(f"- {c}")
    lines.append("")

    lines.append("## Recommendations")
    for r in narrative.recommendations:
        lines.append(f"- {r}")
    lines.append("")

    lines.append("## Key Metrics")
    lines.append("")
    lines.append("| Metric | Score (0-100) | Weight | Detail |")
    lines.append("|---|---|---|---|")
    for mm in result.metrics:
        lines.append(f"| {mm.name.replace('_', ' ').title()} | {mm.score:.0f} | "
                      f"{int(mm.weight * 100)}% | {mm.detail} |")
    lines.append("")

    lines.append("## Project Snapshot")
    s = project.summary
    lines.append(f"- **Stage:** {s.project_stage or 'N/A'}")
    lines.append(f"- **Timeline:** {_fmt_date(s.project_start_date)} → {_fmt_date(s.project_end_date)}")
    lines.append(f"- **Tasks:** {s.completed or 0} completed / {s.in_progress or 0} in progress / "
                  f"{s.not_started or 0} not started / {s.on_hold or 0} on hold "
                  f"(total {len(project.tasks)})")
    lines.append("")

    if project.warnings:
        lines.append("## Data Quality Notes")
        for w in project.warnings:
            lines.append(f"- {w}")
        lines.append("")

    lines.append("---")
    lines.append(f"*Generated automatically by the Project Health Reporting Agent on "
                 f"{datetime.now().strftime('%d %b %Y %H:%M')}.*")
    return "\n".join(lines)


def render_html(project: ParsedProject, result: RagResult, narrative: Narrative) -> str:
    color = RAG_COLOR[result.rag]
    rows = "".join(
        f"<tr><td>{mm.name.replace('_', ' ').title()}</td><td>{mm.score:.0f}</td>"
        f"<td>{int(mm.weight * 100)}%</td><td>{mm.detail}</td></tr>"
        for mm in result.metrics
    )
    disagreement_html = ""
    if not result.status_agrees_with_sheet:
        disagreement_html = (
            f'<div class="callout">⚠️ Independent status differs from source sheet. '
            f'Sheet reads <b>{result.sheet_reported_status}</b>; agent computes <b>{result.rag}</b>. '
            f'See Risk Explanation.</div>'
        )
    positives = "".join(f"<li>{p}</li>" for p in narrative.positive_observations)
    concerns = "".join(f"<li>{c}</li>" for c in narrative.top_concerns)
    recs = "".join(f"<li>{r}</li>" for r in narrative.recommendations)
    warnings_html = ""
    if project.warnings:
        warnings_html = "<h2>Data Quality Notes</h2><ul>" + "".join(f"<li>{w}</li>" for w in project.warnings) + "</ul>"

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">
<title>Weekly Health Report — {project.project_name}</title>
<style>
  @media print {{ body {{ margin: 0.5in; }} }}
  body {{ font-family: Arial, Helvetica, sans-serif; color: #1f2328; max-width: 860px; margin: 40px auto; line-height: 1.5; }}
  h1 {{ font-size: 22px; }}
  h2 {{ font-size: 16px; border-bottom: 1px solid #d0d7de; padding-bottom: 4px; margin-top: 28px; }}
  .status {{ display: inline-block; padding: 4px 14px; border-radius: 14px; color: white; background: {color}; font-weight: bold; }}
  .meta {{ color: #57606a; font-size: 14px; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 13px; }}
  th, td {{ border: 1px solid #d0d7de; padding: 6px 8px; text-align: left; vertical-align: top; }}
  th {{ background: #f6f8fa; }}
  .callout {{ background: #fff8c5; border: 1px solid #d4a72c; padding: 10px 14px; border-radius: 6px; margin: 12px 0; }}
  footer {{ margin-top: 30px; font-size: 12px; color: #8c959f; }}
</style></head>
<body>
  <h1>Weekly Project Health Report — {project.project_name}</h1>
  <p class="meta">Report date: {_fmt_date(result.as_of)} &nbsp;|&nbsp; PM: {project.summary.project_manager or 'N/A'}</p>
  <p><span class="status">{result.rag.upper()}</span>
     &nbsp; composite score {result.composite_score}/100 &nbsp;|&nbsp; confidence {int(result.confidence*100)}%</p>
  {disagreement_html}
  <h2>Executive Summary</h2><p>{narrative.executive_summary}</p>
  <h2>Risk Explanation</h2><p>{narrative.risk_explanation}</p>
  <h2>Positive Observations</h2><ul>{positives}</ul>
  <h2>Top Concerns</h2><ul>{concerns}</ul>
  <h2>Recommendations</h2><ul>{recs}</ul>
  <h2>Key Metrics</h2>
  <table><tr><th>Metric</th><th>Score</th><th>Weight</th><th>Detail</th></tr>{rows}</table>
  {warnings_html}
  <footer>Generated automatically by the Project Health Reporting Agent on {datetime.now().strftime('%d %b %Y %H:%M')}.</footer>
</body></html>"""


def save_weekly_report(project: ParsedProject, result: RagResult, narrative: Narrative, out_dir: str) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    safe_name = "".join(c if c.isalnum() or c in "-_ " else "_" for c in project.project_name).strip().replace(" ", "_")
    date_tag = result.as_of.strftime("%Y-%m-%d") if result.as_of is not None else datetime.now().strftime("%Y-%m-%d")

    md_path = os.path.join(out_dir, f"{safe_name}_{date_tag}.md")
    html_path = os.path.join(out_dir, f"{safe_name}_{date_tag}.html")

    with open(md_path, "w") as f:
        f.write(render_markdown(project, result, narrative))
    with open(html_path, "w") as f:
        f.write(render_html(project, result, narrative))

    return {"markdown": md_path, "html": html_path}
