"""
rag_engine.py
-------------
Deterministic, configurable RAG (Red/Amber/Green) scoring engine.

Design goal (per assignment): do NOT simply copy the sheet's existing
'Schedule Health' or 'RAG' column. Calculate an independent composite
health score from multiple signals, then compare it against whatever
status the sheet already carries and explain any disagreement.

No LLM involved here -- pure, auditable, unit-testable Python.
"""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from excel_parser import ParsedProject

logger = logging.getLogger("rag_engine")

RAG_ORDER = {"Green": 0, "Amber": 1, "Red": 2}

# Keywords that signal a blocker / risk when found in free-text comments
BLOCKER_KEYWORDS = [
    "block", "pending", "delay", "risk", "issue", "escalat", "hold",
    "waiting", "concern", "impact", "not available", "reschedule",
    "overdue", "gap", "dependency", "revisit",
]

DEFAULT_WEIGHTS = {
    "schedule_slippage": 0.30,
    "completion": 0.20,
    "critical_tasks": 0.20,
    "blockers": 0.15,
    "milestones": 0.15,
}


@dataclass
class MetricResult:
    name: str
    score: float          # 0 (worst) - 100 (best)
    weight: float
    detail: str
    facts: dict = field(default_factory=dict)


@dataclass
class RagResult:
    project_name: str
    rag: str
    composite_score: float
    confidence: float
    metrics: list[MetricResult]
    reasons: list[str]
    sheet_reported_status: str | None
    status_agrees_with_sheet: bool
    disagreement_explanation: str | None
    as_of: datetime


class RagEngine:
    def __init__(self, weights: dict | None = None,
                 green_threshold: float = 80.0,
                 amber_threshold: float = 60.0,
                 slippage_days_warn: int = 5,
                 slippage_days_critical: int = 10):
        self.weights = weights or DEFAULT_WEIGHTS
        assert abs(sum(self.weights.values()) - 1.0) < 1e-6, "Weights must sum to 1.0"
        self.green_threshold = green_threshold
        self.amber_threshold = amber_threshold
        self.slippage_days_warn = slippage_days_warn
        self.slippage_days_critical = slippage_days_critical

    # ------------------------------------------------------------------
    def evaluate(self, project: ParsedProject) -> RagResult:
        df = project.tasks
        as_of = project.summary.todays_date or pd.Timestamp.today().normalize()

        metrics = [
            self._score_schedule_slippage(df, as_of),
            self._score_completion(df, project),
            self._score_critical_tasks(df),
            self._score_blockers(project.comments, df),
            self._score_milestones(df),
        ]

        composite = sum(m.score * m.weight for m in metrics)
        rag = self._score_to_rag(composite)

        # Override rules: critical-task or blocker severity can force a worse rating
        rag, override_reason = self._apply_overrides(rag, metrics)

        confidence = self._confidence(df, project, metrics)
        reasons = [m.detail for m in metrics]
        if override_reason:
            reasons.append(override_reason)

        sheet_status = self._normalize_sheet_status(project.summary.schedule_health)
        agrees, disagreement = self._compare_to_sheet(rag, sheet_status, metrics)

        return RagResult(
            project_name=project.project_name,
            rag=rag,
            composite_score=round(composite, 1),
            confidence=confidence,
            metrics=metrics,
            reasons=reasons,
            sheet_reported_status=sheet_status,
            status_agrees_with_sheet=agrees,
            disagreement_explanation=disagreement,
            as_of=as_of,
        )

    # ------------------------------------------------------------------
    # Individual metric scorers
    # ------------------------------------------------------------------

    def _score_schedule_slippage(self, df: pd.DataFrame, as_of) -> MetricResult:
        var_col = "Variance_days" if "Variance_days" in df.columns else None
        variances = df[var_col].dropna() if var_col else pd.Series(dtype=float)

        delayed = variances[variances < 0]
        avg_delay = float(-delayed.mean()) if not delayed.empty else 0.0
        pct_delayed_tasks = (len(delayed) / len(df)) * 100 if len(df) else 0.0

        missed_end_dates = 0
        if "End Date" in df.columns and "Status" in df.columns:
            not_done = df["Status"].fillna("").str.lower() != "completed"
            overdue_mask = df["End Date"].notna() & (df["End Date"] < as_of) & not_done
            missed_end_dates = int(overdue_mask.sum())
        pct_overdue = (missed_end_dates / len(df)) * 100 if len(df) else 0.0

        # Score: penalize on avg delay magnitude, % delayed tasks, % overdue-open tasks
        score = 100.0
        score -= min(avg_delay * 3, 40)
        score -= min(pct_delayed_tasks * 0.8, 30)
        score -= min(pct_overdue * 1.0, 30)
        score = max(0.0, min(100.0, score))

        detail = (
            f"Schedule slippage: {len(delayed)} of {len(df)} tasks ({pct_delayed_tasks:.0f}%) show negative "
            f"variance, averaging {avg_delay:.1f} days late where delayed; {missed_end_dates} open tasks "
            f"({pct_overdue:.0f}%) are already past their planned end date."
        )
        return MetricResult("schedule_slippage", score, self.weights["schedule_slippage"], detail, {
            "avg_delay_days": round(avg_delay, 1),
            "pct_delayed_tasks": round(pct_delayed_tasks, 1),
            "missed_end_dates": missed_end_dates,
            "pct_overdue_open": round(pct_overdue, 1),
        })

    def _score_completion(self, df: pd.DataFrame, project: ParsedProject) -> MetricResult:
        if "% Complete" in df.columns and df["% Complete"].notna().any():
            avg_complete = float(df["% Complete"].dropna().mean()) * 100
        elif project.summary.pct_complete is not None:
            avg_complete = project.summary.pct_complete * 100
        else:
            avg_complete = 50.0  # neutral fallback when no data at all

        # Time-elapsed vs completion gap, when start/end dates are available
        elapsed_gap_penalty = 0.0
        gap_note = ""
        start, end, today = project.summary.project_start_date, project.summary.project_end_date, project.summary.todays_date
        if start is not None and end is not None and today is not None and end > start:
            total_days = (end - start).days
            elapsed_days = max(0, min((today - start).days, total_days))
            expected_pct = (elapsed_days / total_days) * 100 if total_days else 0
            gap = expected_pct - avg_complete
            if gap > 0:
                elapsed_gap_penalty = min(gap * 0.8, 35)
                gap_note = f" Timeline is {expected_pct:.0f}% elapsed vs {avg_complete:.0f}% complete (a {gap:.0f}-point gap)."

        score = avg_complete - elapsed_gap_penalty
        score = max(0.0, min(100.0, score))

        detail = f"Completion: {avg_complete:.0f}% of tasks marked complete." + gap_note
        return MetricResult("completion", score, self.weights["completion"], detail, {
            "avg_pct_complete": round(avg_complete, 1),
            "elapsed_gap_penalty": round(elapsed_gap_penalty, 1),
        })

    def _score_critical_tasks(self, df: pd.DataFrame) -> MetricResult:
        if "Critical?" not in df.columns or not df["Critical?"].any():
            return MetricResult("critical_tasks", 90.0, self.weights["critical_tasks"],
                                 "No tasks flagged Critical; treating critical-path risk as low.",
                                 {"critical_count": 0})

        crit = df[df["Critical?"] == True]
        n_crit = len(crit)
        crit_red = 0
        if "Schedule Health" in crit.columns:
            crit_red = int((crit["Schedule Health"].fillna("").str.lower() == "red").sum())
        crit_not_done = 0
        if "Status" in crit.columns:
            crit_not_done = int((crit["Status"].fillna("").str.lower() != "completed").sum())
        crit_delayed = 0
        if "Variance_days" in crit.columns:
            crit_delayed = int((crit["Variance_days"].fillna(0) < 0).sum())

        pct_red = (crit_red / n_crit) * 100 if n_crit else 0
        pct_delayed = (crit_delayed / n_crit) * 100 if n_crit else 0

        score = 100.0
        score -= min(pct_red * 0.9, 60)
        score -= min(pct_delayed * 0.4, 25)
        score = max(0.0, min(100.0, score))

        detail = (
            f"Critical-path tasks: {n_crit} flagged critical; {crit_red} ({pct_red:.0f}%) already rated Red "
            f"internally and {crit_delayed} ({pct_delayed:.0f}%) show schedule slippage; {crit_not_done} remain open."
        )
        return MetricResult("critical_tasks", score, self.weights["critical_tasks"], detail, {
            "critical_count": n_crit, "critical_red": crit_red, "critical_delayed": crit_delayed,
            "critical_open": crit_not_done,
        })

    def _score_blockers(self, comments: list, df: pd.DataFrame) -> MetricResult:
        texts = [c.text for c in comments]
        # also fold in any populated 'Status Comment' cells on tasks, if present
        if "Status Comment" in df.columns:
            texts += [t for t in df["Status Comment"].dropna().tolist()]

        if not texts:
            return MetricResult("blockers", 85.0, self.weights["blockers"],
                                 "No status comments were logged for this period; blocker visibility is limited.",
                                 {"comment_count": 0, "flagged_count": 0})

        flagged = [t for t in texts if re.search("|".join(BLOCKER_KEYWORDS), t.lower())]
        pct_flagged = (len(flagged) / len(texts)) * 100

        score = 100.0 - min(pct_flagged * 0.8, 55)
        score = max(0.0, min(100.0, score))

        sample = "; ".join(flagged[:2]) if flagged else ""
        detail = (
            f"Stakeholder comments: {len(texts)} logged, {len(flagged)} ({pct_flagged:.0f}%) reference "
            f"blockers, pending approvals, or risk."
        )
        if sample:
            detail += f" Representative: \"{sample[:140]}\""
        return MetricResult("blockers", score, self.weights["blockers"], detail, {
            "comment_count": len(texts), "flagged_count": len(flagged),
        })

    def _score_milestones(self, df: pd.DataFrame) -> MetricResult:
        if "Phase/Milestone" not in df.columns:
            return MetricResult("milestones", 80.0, self.weights["milestones"],
                                 "No milestone/phase field present in source data.", {"milestone_count": 0})

        milestone_rows = df[df["Phase/Milestone"].notna()]
        if milestone_rows.empty:
            return MetricResult("milestones", 80.0, self.weights["milestones"],
                                 "No milestone rows identified in the plan.", {"milestone_count": 0})

        n = len(milestone_rows)
        completed = 0
        if "Status" in milestone_rows.columns:
            completed = int((milestone_rows["Status"].fillna("").str.lower() == "completed").sum())
        pending_overdue = 0
        if "Variance_days" in milestone_rows.columns:
            pending_overdue = int((milestone_rows["Variance_days"].fillna(0) < 0).sum())

        pct_complete = (completed / n) * 100 if n else 0
        pct_overdue = (pending_overdue / n) * 100 if n else 0

        score = pct_complete - min(pct_overdue * 0.5, 30)
        score = max(0.0, min(100.0, score))

        detail = (
            f"Milestones: {completed}/{n} phase-level milestones completed ({pct_complete:.0f}%); "
            f"{pending_overdue} show schedule variance."
        )
        return MetricResult("milestones", score, self.weights["milestones"], detail, {
            "milestone_count": n, "milestones_completed": completed, "milestones_overdue": pending_overdue,
        })

    # ------------------------------------------------------------------
    def _score_to_rag(self, composite: float) -> str:
        if composite >= self.green_threshold:
            return "Green"
        if composite >= self.amber_threshold:
            return "Amber"
        return "Red"

    def _apply_overrides(self, rag: str, metrics: list[MetricResult]) -> tuple[str, str | None]:
        """Guardrails: severe critical-task or blocker signals cap the rating,
        even if the weighted composite would otherwise round up."""
        crit = next(m for m in metrics if m.name == "critical_tasks")
        blockers = next(m for m in metrics if m.name == "blockers")

        worst = rag
        reason = None
        if crit.facts.get("critical_count", 0) > 0:
            pct_red = crit.facts["critical_red"] / crit.facts["critical_count"] * 100
            if pct_red >= 30 and RAG_ORDER[rag] < RAG_ORDER["Red"]:
                worst = "Red"
                reason = (f"Override: {pct_red:.0f}% of critical-path tasks are internally rated Red, "
                          f"which caps the overall status at Red regardless of the composite score.")
            elif pct_red >= 10 and RAG_ORDER[rag] < RAG_ORDER["Amber"]:
                worst = "Amber"
                reason = (f"Override: {pct_red:.0f}% of critical-path tasks are Red, capping the overall "
                          f"status at Amber even though other signals are healthy.")

        if blockers.facts.get("comment_count", 0) >= 3:
            pct_flag = blockers.facts["flagged_count"] / blockers.facts["comment_count"] * 100
            if pct_flag >= 60 and RAG_ORDER[worst] < RAG_ORDER["Amber"]:
                worst = "Amber"
                reason = (reason or "") + (
                    f" Override: {pct_flag:.0f}% of logged comments flag blockers or pending items, "
                    f"which caps the status at Amber."
                )

        return worst, reason.strip() if reason else None

    def _confidence(self, df: pd.DataFrame, project: ParsedProject, metrics: list[MetricResult]) -> float:
        """0-1 confidence based on data completeness feeding the score."""
        signals_present = 0
        total_signals = 5
        if "Variance_days" in df.columns and df["Variance_days"].notna().any():
            signals_present += 1
        if "% Complete" in df.columns and df["% Complete"].notna().any():
            signals_present += 1
        if "Critical?" in df.columns and df["Critical?"].any():
            signals_present += 1
        if project.comments or ("Status Comment" in df.columns and df["Status Comment"].notna().any()):
            signals_present += 1
        if "Phase/Milestone" in df.columns and df["Phase/Milestone"].notna().any():
            signals_present += 1
        base = signals_present / total_signals
        penalty = 0.1 if len(project.warnings) > 0 else 0.0
        return round(max(0.3, base - penalty), 2)

    def _normalize_sheet_status(self, value) -> str | None:
        if value is None:
            return None
        v = str(value).strip().lower()
        if v.startswith("green"):
            return "Green"
        if v.startswith("yellow") or v.startswith("amber"):
            return "Amber"
        if v.startswith("red"):
            return "Red"
        return None

    def _compare_to_sheet(self, computed_rag: str, sheet_status: str | None,
                           metrics: list[MetricResult]) -> tuple[bool, str | None]:
        if sheet_status is None:
            return True, None
        if sheet_status == computed_rag:
            return True, None

        crit = next(m for m in metrics if m.name == "critical_tasks")
        slip = next(m for m in metrics if m.name == "schedule_slippage")

        direction = "more severe" if RAG_ORDER[computed_rag] > RAG_ORDER[sheet_status] else "less severe"
        explanation = (
            f"The project-level Schedule Health field in the source sheet reads '{sheet_status}', but the "
            f"independently computed status is '{computed_rag}' -- {direction}. This is a composite view: "
            f"individual task-level health can look fine in isolation while critical-path delays "
            f"({slip.facts.get('missed_end_dates', 0)} overdue open tasks) and critical-task risk "
            f"({crit.facts.get('critical_red', 0)} of {crit.facts.get('critical_count', 0)} critical tasks "
            f"rated Red) pull the overall picture in a different direction than any single field suggests."
        )
        return False, explanation
