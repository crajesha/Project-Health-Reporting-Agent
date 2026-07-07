"""
reasoning.py
------------
Turns the deterministic RagResult into plain-English narrative:
executive summary, risk explanation, positive observations, top
concerns, and recommendations -- written in the voice of a senior
delivery manager, not a generic AI disclaimer-fest.

Implementation note: this is a template-driven natural-language
generator, not a call to an external LLM. It is deterministic (so
outputs are reproducible and auditable) and has zero external
dependency / API-key requirement, which keeps the agent runnable
anywhere. See README "Design Decisions" for why this tradeoff was
made, and how to swap in a hosted LLM (e.g. Claude or Gemini) behind
the same interface if richer prose is wanted later.
"""

from __future__ import annotations

from dataclasses import dataclass

from excel_parser import ParsedProject
from rag_engine import RagResult


@dataclass
class Narrative:
    executive_summary: str
    risk_explanation: str
    positive_observations: list[str]
    top_concerns: list[str]
    recommendations: list[str]


def _metric(result: RagResult, name: str):
    return next(m for m in result.metrics if m.name == name)


def build_narrative(project: ParsedProject, result: RagResult) -> Narrative:
    slip = _metric(result, "schedule_slippage")
    comp = _metric(result, "completion")
    crit = _metric(result, "critical_tasks")
    block = _metric(result, "blockers")
    mile = _metric(result, "milestones")

    pm = project.summary.project_manager or "the project manager"
    stage = project.summary.project_stage or "the current phase"

    # ---------------- Executive summary ----------------
    tone = {
        "Green": "is tracking well",
        "Amber": "needs active management attention",
        "Red": "requires immediate executive intervention",
    }[result.rag]

    exec_summary = (
        f"{project.project_name} is rated {result.rag} ({result.composite_score}/100) as of "
        f"{result.as_of.strftime('%d %b %Y') if result.as_of is not None else 'the latest update'} "
        f"and {tone}. The project is currently in {stage}, {comp.facts['avg_pct_complete']:.0f}% complete "
        f"under {pm}."
    )
    if result.disagreement_explanation:
        exec_summary += (
            f" Note: this differs from the sheet's self-reported Schedule Health "
            f"('{result.sheet_reported_status}') -- see risk explanation below."
        )

    # ---------------- Risk explanation ----------------
    risk_bits = []
    if slip.facts["missed_end_dates"] > 0:
        risk_bits.append(
            f"{slip.facts['missed_end_dates']} tasks are past their planned end date "
            f"({slip.facts['pct_overdue_open']:.0f}% of the plan), averaging "
            f"{slip.facts['avg_delay_days']:.0f} days of slippage where delays exist."
        )
    if crit.facts.get("critical_count", 0) > 0 and crit.facts.get("critical_red", 0) > 0:
        risk_bits.append(
            f"Of {crit.facts['critical_count']} critical-path tasks, {crit.facts['critical_red']} are "
            f"already internally rated Red -- these sit on the path that most directly threatens the "
            f"end date."
        )
    elif crit.facts.get("critical_delayed", 0) > 0:
        risk_bits.append(
            f"{crit.facts['critical_delayed']} of {crit.facts['critical_count']} critical-path tasks "
            f"show negative schedule variance, which is the leading indicator worth watching even though "
            f"none are formally flagged Red yet."
        )
    if block.facts.get("flagged_count", 0) > 0:
        risk_bits.append(
            f"{block.facts['flagged_count']} of {block.facts['comment_count']} recent stakeholder comments "
            f"reference blockers, pending approvals, or dependencies -- a leading indicator that often "
            f"precedes schedule slippage in the following reporting cycle."
        )
    if result.disagreement_explanation:
        risk_bits.append(result.disagreement_explanation)

    risk_explanation = " ".join(risk_bits) if risk_bits else (
        "No material risk signals were detected in this cycle's data -- schedule variance, critical-task "
        "health, and stakeholder comments are all within normal range."
    )

    # ---------------- Positive observations ----------------
    positives = []
    if comp.facts["avg_pct_complete"] >= 50:
        positives.append(f"Overall completion stands at {comp.facts['avg_pct_complete']:.0f}%, "
                          f"showing steady forward progress.")
    if slip.facts["pct_delayed_tasks"] < 10:
        positives.append(f"Only {slip.facts['pct_delayed_tasks']:.0f}% of tasks show negative variance, "
                          f"indicating the bulk of the plan is on schedule.")
    if mile.facts.get("milestone_count", 0) > 0 and mile.facts["milestones_completed"] > 0:
        positives.append(f"{mile.facts['milestones_completed']} of {mile.facts['milestone_count']} "
                          f"tracked milestones are already complete.")
    if block.facts.get("comment_count", 0) > 0 and block.facts["flagged_count"] / max(block.facts["comment_count"], 1) < 0.3:
        positives.append("Stakeholder sentiment in recent comments is largely constructive, with few "
                          "escalation-level flags.")
    if not positives:
        positives.append("Team engagement remains active based on comment cadence, even though "
                          "quantitative signals are weak this cycle.")

    # ---------------- Top concerns ----------------
    concerns = []
    if slip.facts["missed_end_dates"] > 0:
        concerns.append(f"{slip.facts['missed_end_dates']} open tasks past their planned end date.")
    if crit.facts.get("critical_red", 0) > 0:
        concerns.append(f"{crit.facts['critical_red']} critical-path tasks rated Red.")
    if crit.facts.get("critical_delayed", 0) > 0 and crit.facts.get("critical_red", 0) == 0:
        concerns.append(f"{crit.facts['critical_delayed']} critical-path tasks trending behind schedule.")
    if block.facts.get("flagged_count", 0) > 0:
        concerns.append(f"{block.facts['flagged_count']} stakeholder comments flag pending approvals or blockers.")
    if comp.facts.get("elapsed_gap_penalty", 0) > 5:
        concerns.append("Completion is lagging the elapsed timeline, widening the delivery gap.")
    if not concerns:
        concerns.append("No significant concerns identified this cycle; maintain current cadence.")

    # ---------------- Recommendations ----------------
    recs = []
    if result.rag == "Red":
        recs.append("Escalate to a steering-committee review within the next 5 business days; the "
                     "current trajectory will miss the committed end date without intervention.")
    elif result.rag == "Amber":
        recs.append("Schedule a focused checkpoint with the project manager this week to agree a "
                     "recovery plan for the critical-path items called out above.")
    else:
        recs.append("Maintain current governance cadence; no escalation required this cycle.")

    if crit.facts.get("critical_red", 0) > 0 or crit.facts.get("critical_delayed", 0) > 0:
        recs.append("Re-baseline or re-sequence the affected critical-path tasks and confirm "
                     "resourcing with the owning team leads.")
    if block.facts.get("flagged_count", 0) > 0:
        recs.append("Follow up directly with the stakeholders behind the flagged comments to close "
                     "out pending approvals before they convert into schedule slippage.")
    if slip.facts["pct_overdue_open"] > 15:
        recs.append("Run a scope/priority triage on overdue tasks to distinguish true blockers from "
                     "tasks that can simply be re-dated.")

    return Narrative(
        executive_summary=exec_summary,
        risk_explanation=risk_explanation,
        positive_observations=positives[:4],
        top_concerns=concerns[:4],
        recommendations=recs[:4],
    )
