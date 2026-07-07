"""
excel_parser.py
----------------
Enterprise-grade parser for Project Health Reporting Agent.

Reads a project plan workbook that follows the observed convention:
  - One "project plan" sheet (name varies per project) with task-level rows
  - A "Comments" sheet with free-text status comments (row ref, text, author, timestamp)
  - A "Summary" sheet with project-level key/value rollup metrics

Handles messy real-world data: #UNPARSEABLE tokens, blank cells, NaN,
mixed date formats, and inconsistent column presence across files.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("excel_parser")

# Tokens that appear in the source sheets but mean "no usable value"
NULL_TOKENS = {"#UNPARSEABLE", "#N/A", "N/A", "NA", "", "NONE", "NULL", "-"}


def _is_null_token(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    if isinstance(value, str) and value.strip().upper() in NULL_TOKENS:
        return True
    return False


def clean_value(value):
    """Normalize a single cell value: strips junk tokens down to None."""
    if _is_null_token(value):
        return None
    if isinstance(value, str):
        return value.strip()
    return value


def parse_variance_days(value) -> Optional[int]:
    """Convert values like '-8d', '0', '15d' into signed integer days."""
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip().lower().replace("d", "")
    try:
        return int(text)
    except ValueError:
        return None


def to_datetime(value) -> Optional[pd.Timestamp]:
    value = clean_value(value)
    if value is None:
        return None
    try:
        ts = pd.to_datetime(value, errors="coerce")
        return None if pd.isna(ts) else ts
    except Exception:
        return None


def normalize_columns(columns) -> list[str]:
    """Trim whitespace/odd characters from column headers without renaming meaning."""
    cleaned = []
    for c in columns:
        c = str(c).strip()
        c = c.replace("Critical ?", "Critical?")
        cleaned.append(c)
    return cleaned


@dataclass
class CommentEntry:
    row_ref: Optional[str]
    text: str
    author: Optional[str]
    timestamp: Optional[pd.Timestamp]


@dataclass
class ProjectSummary:
    project_manager: Optional[str] = None
    project_start_date: Optional[pd.Timestamp] = None
    project_end_date: Optional[pd.Timestamp] = None
    not_started: Optional[int] = None
    in_progress: Optional[int] = None
    completed: Optional[int] = None
    on_hold: Optional[int] = None
    at_risk: Optional[str] = None
    project_stage: Optional[str] = None
    pct_complete: Optional[float] = None
    schedule_health: Optional[str] = None
    todays_date: Optional[pd.Timestamp] = None
    project_status: Optional[str] = None
    raw: dict = field(default_factory=dict)


@dataclass
class ParsedProject:
    source_file: str
    project_name: str
    tasks: pd.DataFrame
    comments: list[CommentEntry]
    summary: ProjectSummary
    warnings: list[str] = field(default_factory=list)


SUMMARY_KEY_MAP = {
    "project manager": "project_manager",
    "project start date": "project_start_date",
    "project end date": "project_end_date",
    "not started": "not_started",
    "in progress": "in_progress",
    "completed": "completed",
    "on hold": "on_hold",
    "at risk": "at_risk",
    "project stage": "project_stage",
    "% complete": "pct_complete",
    "schedule health": "schedule_health",
    "today's date": "todays_date",
    "project status": "project_status",
}

DATE_FIELDS = {"project_start_date", "project_end_date", "todays_date"}
INT_FIELDS = {"not_started", "in_progress", "completed", "on_hold"}


class ExcelParser:
    """Parses one or more project-plan workbooks into ParsedProject objects."""

    def parse_file(self, filepath: str) -> ParsedProject:
        logger.info("Parsing workbook: %s", filepath)
        warnings: list[str] = []

        try:
            xl = pd.ExcelFile(filepath)
        except Exception as exc:
            logger.error("Failed to open workbook %s: %s", filepath, exc)
            raise

        task_sheet_name = self._detect_task_sheet(xl, warnings)
        tasks_df = self._parse_task_sheet(xl, task_sheet_name, warnings)
        comments = self._parse_comments_sheet(xl, warnings)
        summary = self._parse_summary_sheet(xl, warnings)

        project_name = self._infer_project_name(tasks_df, task_sheet_name, filepath)

        parsed = ParsedProject(
            source_file=filepath,
            project_name=project_name,
            tasks=tasks_df,
            comments=comments,
            summary=summary,
            warnings=warnings,
        )
        logger.info(
            "Parsed '%s': %d tasks, %d comments, %d warnings",
            project_name, len(tasks_df), len(comments), len(warnings),
        )
        return parsed

    def parse_files(self, filepaths: list[str]) -> list[ParsedProject]:
        results = []
        for fp in filepaths:
            try:
                results.append(self.parse_file(fp))
            except Exception as exc:
                logger.error("Skipping file %s due to error: %s", fp, exc)
        return results

    # ------------------------------------------------------------------
    # Sheet detection
    # ------------------------------------------------------------------

    def _detect_task_sheet(self, xl: pd.ExcelFile, warnings: list[str]) -> str:
        """The task-level sheet is whichever sheet is NOT 'Comments' or 'Summary',
        preferring the one with the most rows/columns if several candidates exist."""
        candidates = [s for s in xl.sheet_names if s.strip().lower() not in ("comments", "summary")]
        if not candidates:
            warnings.append("No task sheet found distinct from Comments/Summary; defaulting to first sheet.")
            return xl.sheet_names[0]
        if len(candidates) == 1:
            return candidates[0]
        # multiple candidates: pick the widest/tallest sheet
        best, best_size = candidates[0], -1
        for name in candidates:
            df = xl.parse(name, nrows=5)
            size = df.shape[1]
            if size > best_size:
                best, best_size = name, size
        return best

    # ------------------------------------------------------------------
    # Task sheet
    # ------------------------------------------------------------------

    def _parse_task_sheet(self, xl: pd.ExcelFile, sheet_name: str, warnings: list[str]) -> pd.DataFrame:
        try:
            df = xl.parse(sheet_name)
        except Exception as exc:
            warnings.append(f"Could not read task sheet '{sheet_name}': {exc}")
            return pd.DataFrame()

        df.columns = normalize_columns(df.columns)

        # Clean every object column of junk tokens
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].map(clean_value)

        # Parse known date columns
        for col in ["Start Date", "End Date", "Baseline Start", "Baseline Finish",
                    "Baseline Start Date", "Baseline End Date", "Start", "Finish",
                    "Baseline Start2", "Baseline Finish2"]:
            if col in df.columns:
                df[col] = df[col].map(to_datetime)

        # Parse variance-style columns into signed integer days
        for col in ["Variance", "Variance2"]:
            if col in df.columns:
                df[col + "_days"] = df[col].map(parse_variance_days)

        # Normalize % Complete to a 0-1 float
        if "% Complete" in df.columns:
            df["% Complete"] = pd.to_numeric(df["% Complete"], errors="coerce")
            if df["% Complete"].max(skipna=True) is not None and df["% Complete"].max(skipna=True) > 1.5:
                df["% Complete"] = df["% Complete"] / 100.0

        # Normalize Critical? flag to boolean (handles NaN explicitly --
        # bool(float('nan')) is True in Python, so a naive cast is wrong here)
        if "Critical?" in df.columns:
            df["Critical?"] = df["Critical?"].apply(
                lambda v: bool(pd.notna(v)) and v not in (0, "0", False, "False", "N")
            )
        else:
            df["Critical?"] = False

        # Drop fully blank rows
        before = len(df)
        df = df.dropna(how="all")
        dropped = before - len(df)
        if dropped:
            warnings.append(f"Dropped {dropped} fully blank rows from task sheet.")

        return df.reset_index(drop=True)

    # ------------------------------------------------------------------
    # Comments sheet
    # ------------------------------------------------------------------

    def _parse_comments_sheet(self, xl: pd.ExcelFile, warnings: list[str]) -> list[CommentEntry]:
        name = next((s for s in xl.sheet_names if s.strip().lower() == "comments"), None)
        if name is None:
            warnings.append("No Comments sheet present.")
            return []
        try:
            raw = xl.parse(name, header=None)
        except Exception as exc:
            warnings.append(f"Could not read Comments sheet: {exc}")
            return []

        entries: list[CommentEntry] = []
        for _, row in raw.iterrows():
            values = list(row)
            if len(values) < 2:
                continue
            row_ref, text = clean_value(values[0]), clean_value(values[1])
            author = clean_value(values[2]) if len(values) > 2 else None
            ts_raw = clean_value(values[3]) if len(values) > 3 else None
            if text is None:
                continue  # skip spacer rows
            entries.append(CommentEntry(
                row_ref=row_ref,
                text=str(text),
                author=author,
                timestamp=to_datetime(ts_raw),
            ))
        return entries

    # ------------------------------------------------------------------
    # Summary sheet
    # ------------------------------------------------------------------

    def _parse_summary_sheet(self, xl: pd.ExcelFile, warnings: list[str]) -> ProjectSummary:
        name = next((s for s in xl.sheet_names if s.strip().lower() == "summary"), None)
        if name is None:
            warnings.append("No Summary sheet present.")
            return ProjectSummary()
        try:
            raw = xl.parse(name, header=0)
        except Exception as exc:
            warnings.append(f"Could not read Summary sheet: {exc}")
            return ProjectSummary()

        if raw.shape[1] < 2:
            warnings.append("Summary sheet does not have key/value columns as expected.")
            return ProjectSummary()

        kv = {}
        for _, row in raw.iterrows():
            key = clean_value(row.iloc[0])
            val = clean_value(row.iloc[1])
            if key is None:
                continue
            kv[str(key).strip().lower()] = val

        summary = ProjectSummary(raw=kv)
        for raw_key, field_name in SUMMARY_KEY_MAP.items():
            if raw_key in kv:
                val = kv[raw_key]
                if field_name in DATE_FIELDS:
                    val = to_datetime(val)
                elif field_name in INT_FIELDS:
                    try:
                        val = int(val) if val is not None else None
                    except (ValueError, TypeError):
                        val = None
                elif field_name == "pct_complete":
                    try:
                        val = float(val)
                        if val > 1.5:
                            val = val / 100.0
                    except (ValueError, TypeError):
                        val = None
                setattr(summary, field_name, val)
        return summary

    # ------------------------------------------------------------------
    def _infer_project_name(self, tasks_df: pd.DataFrame, sheet_name: str, filepath: str) -> str:
        if "Project Name" in tasks_df.columns:
            names = tasks_df["Project Name"].dropna()
            if not names.empty:
                return str(names.iloc[0])
        # fall back to sheet name if it looks like a project name (not generic)
        if sheet_name.strip().lower() not in ("project plan", "sheet1", "tasks"):
            return sheet_name.strip()
        # fall back to filename
        import os
        return os.path.splitext(os.path.basename(filepath))[0]
