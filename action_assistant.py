"""
action_assistant.py
--------------------
"Executive Action Assistant" (bonus feature).

Turns a delayed / at-risk task into a ready-to-review email DRAFT for
a VP to send to the relevant owner. This module NEVER sends email --
it only produces suggested text that a human reviews, edits, and
sends themselves via their own mail client. See README for rationale.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import pandas as pd

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")


@dataclass
class ActionItem:
    task_name: str
    owner: str
    owner_email: str | None
    days_late: int
    reason: str
    suggested_subject: str
    suggested_body: str


def _extract_email(text) -> str | None:
    if not isinstance(text, str):
        return None
    m = EMAIL_RE.search(text)
    return m.group(0) if m else None


def _first_name(owner: str) -> str:
    if not owner or owner.strip().lower() == "the task owner":
        return "there"
    if "@" in owner:
        local = owner.split("@")[0]
        return local.split(".")[0].capitalize()
    return owner.split()[0]


def build_action_items(project_name: str, tasks: pd.DataFrame, top_n: int = 5) -> list[ActionItem]:
    """Identify the most at-risk open tasks and draft a follow-up email for each."""
    if tasks.empty or "Variance_days" not in tasks.columns:
        return []

    candidates = tasks.copy()
    if "Status" in candidates.columns:
        candidates = candidates[candidates["Status"].fillna("").str.lower() != "completed"]
    candidates = candidates[candidates["Variance_days"].fillna(0) < 0]
    if candidates.empty:
        return []

    candidates = candidates.sort_values("Variance_days").head(top_n)

    items = []
    for _, row in candidates.iterrows():
        task_name = str(row.get("Task Name") or "Untitled task")
        days_late = int(-row["Variance_days"]) if pd.notna(row["Variance_days"]) else 0

        assigned_to = row.get("Assigned To")
        owner_col = row.get("Owner")
        owner_display = "the task owner"
        for candidate in (assigned_to, owner_col):
            if isinstance(candidate, str) and candidate.strip():
                owner_display = candidate.strip()
                break
        owner_email = _extract_email(owner_display)

        comment = row.get("Status Comment") or row.get("Comments")
        if isinstance(comment, float):  # NaN
            comment = None
        reason = f"{task_name} is delayed by {days_late} day(s)"
        if isinstance(comment, str) and comment.strip():
            reason += f" ({comment.strip()})"
        else:
            reason += " based on current schedule variance"

        subject = f"Action Required — {task_name} Delay"
        greeting = _first_name(owner_display)
        body = (
            f"Hi {greeting},\n\n"
            f"The \"{task_name}\" task for the {project_name} project is currently delayed by "
            f"{days_late} day(s){' due to ' + comment.strip() if isinstance(comment, str) and comment.strip() else ''}. "
            f"To avoid impacting downstream activities, could you review this and share an updated "
            f"timeline by end of this week?\n\n"
            f"Please let us know if additional support is needed to get this back on track.\n\n"
            f"Regards,\nProject Health Reporting Agent (on behalf of Professional Services)"
        )

        items.append(ActionItem(
            task_name=task_name,
            owner=owner_display,
            owner_email=owner_email,
            days_late=days_late,
            reason=reason,
            suggested_subject=subject,
            suggested_body=body,
        ))
    return items


def render_action_items_markdown(project_name: str, items: list[ActionItem]) -> str:
    if not items:
        return f"No overdue open tasks requiring follow-up were identified for {project_name} this cycle."

    lines = [f"## Suggested Follow-ups — {project_name}", "",
              "*Drafts only. Review and send manually via your own mail client — this tool does not send email.*", ""]
    for i, item in enumerate(items, 1):
        lines.append(f"### {i}. {item.task_name} — {item.days_late} day(s) late")
        lines.append(f"**To:** {item.owner}" + (f" ({item.owner_email})" if item.owner_email else ""))
        lines.append(f"**Subject:** {item.suggested_subject}")
        lines.append("")
        lines.append("```")
        lines.append(item.suggested_body)
        lines.append("```")
        lines.append("")
    return "\n".join(lines)
