"""
dashboard.py
------------
Streamlit dashboard for the Project Health Reporting Agent.

Run with:
    streamlit run src/dashboard.py
"""

from __future__ import annotations

import io
import os
import sys

import pandas as pd
import plotly.express as px
import streamlit as st

sys.path.insert(0, os.path.dirname(__file__))
from excel_parser import ExcelParser
from rag_engine import RagEngine
from reasoning import build_narrative
from action_assistant import build_action_items
from pptx_builder import build_executive_pptx

st.set_page_config(page_title="Project Health Dashboard", layout="wide")

RAG_COLOR = {"Green": "#1a7f37", "Amber": "#b98400", "Red": "#cf222e"}

st.title("📊 Project Health Reporting — Dashboard")
st.caption("Upload one or more project plan workbooks to get an independent RAG assessment.")

uploaded = st.sidebar.file_uploader("Upload Excel project plan(s)", type=["xlsx"], accept_multiple_files=True)

if not uploaded:
    st.info("Upload one or more `.xlsx` project plans in the sidebar to begin.")
    st.stop()

parser = ExcelParser()
engine = RagEngine()

results = []
for f in uploaded:
    tmp_path = os.path.join("/tmp", f.name)
    with open(tmp_path, "wb") as out:
        out.write(f.getbuffer())
    project = parser.parse_file(tmp_path)
    rag_result = engine.evaluate(project)
    narrative = build_narrative(project, rag_result)
    results.append((project, rag_result, narrative))

# ---------------- Portfolio overview ----------------
st.subheader("Portfolio Overview")

portfolio_data = []
for project, rag_result, _ in results:
    m = {mm.name: mm for mm in rag_result.metrics}
    portfolio_data.append({
        "name": project.project_name,
        "pm": project.summary.project_manager,
        "rag": rag_result.rag,
        "score": rag_result.composite_score,
        "confidence": rag_result.confidence,
        "sheet_status": rag_result.sheet_reported_status,
        "agrees": rag_result.status_agrees_with_sheet,
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
    })

deck_col1, deck_col2 = st.columns([3, 1])
with deck_col2:
    pptx_buffer = io.BytesIO()
    build_executive_pptx(portfolio_data, pptx_buffer)
    pptx_buffer.seek(0)
    st.download_button(
        "📊 Download Executive Presentation",
        data=pptx_buffer,
        file_name="Executive_Project_Health_Review.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
    )

cols = st.columns(len(results)) if len(results) <= 4 else st.columns(4)
for i, (project, rag_result, _) in enumerate(results):
    col = cols[i % len(cols)]
    with col:
        color = RAG_COLOR[rag_result.rag]
        st.markdown(
            f"<div style='border:1px solid #ddd;border-radius:10px;padding:14px;'>"
            f"<b>{project.project_name}</b><br>"
            f"<span style='background:{color};color:white;border-radius:12px;padding:2px 10px;font-weight:bold;'>"
            f"{rag_result.rag}</span> &nbsp; {rag_result.composite_score}/100<br>"
            f"<span style='color:#666;font-size:12px;'>Confidence {int(rag_result.confidence*100)}%</span>"
            f"</div>", unsafe_allow_html=True,
        )

# ---------------- Charts ----------------
st.subheader("Portfolio Charts")
chart_col1, chart_col2 = st.columns(2)

rag_counts = pd.Series([r.rag for _, r, _ in results]).value_counts().reindex(["Green", "Amber", "Red"]).fillna(0)
with chart_col1:
    fig = px.pie(names=rag_counts.index, values=rag_counts.values,
                 color=rag_counts.index, color_discrete_map=RAG_COLOR, title="RAG Distribution")
    st.plotly_chart(fig, use_container_width=True)

score_df = pd.DataFrame({
    "Project": [p.project_name for p, _, _ in results],
    "Score": [r.composite_score for _, r, _ in results],
    "RAG": [r.rag for _, r, _ in results],
})
with chart_col2:
    fig2 = px.bar(score_df, x="Score", y="Project", color="RAG", orientation="h",
                   color_discrete_map=RAG_COLOR, range_x=[0, 100], title="Composite Health Score")
    st.plotly_chart(fig2, use_container_width=True)

# ---------------- Per-project detail ----------------
st.subheader("Project Detail")
tabs = st.tabs([p.project_name for p, _, _ in results])
for tab, (project, rag_result, narrative) in zip(tabs, results):
    with tab:
        st.markdown(f"### {rag_result.rag} — {rag_result.composite_score}/100")
        if not rag_result.status_agrees_with_sheet:
            st.warning(
                f"Independent status differs from source sheet "
                f"(sheet reads **{rag_result.sheet_reported_status}**)."
            )
        st.write("**Executive Summary**")
        st.write(narrative.executive_summary)
        st.write("**Risk Explanation**")
        st.write(narrative.risk_explanation)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.write("**Positive Observations**")
            for p in narrative.positive_observations:
                st.markdown(f"- {p}")
        with c2:
            st.write("**Top Concerns**")
            for c in narrative.top_concerns:
                st.markdown(f"- {c}")
        with c3:
            st.write("**Recommendations**")
            for r in narrative.recommendations:
                st.markdown(f"- {r}")

        st.write("**Key Metrics**")
        metrics_df = pd.DataFrame([
            {"Metric": m.name.replace("_", " ").title(), "Score": m.score, "Weight": f"{int(m.weight*100)}%", "Detail": m.detail}
            for m in rag_result.metrics
        ])
        st.dataframe(metrics_df, use_container_width=True, hide_index=True)

        st.write("**⭐ Executive Action Assistant — Suggested Follow-ups (drafts only, not sent)**")
        action_items = build_action_items(project.project_name, project.tasks)
        if not action_items:
            st.caption("No overdue open tasks requiring follow-up this cycle.")
        for item in action_items:
            with st.expander(f"{item.task_name} — {item.days_late} day(s) late — To: {item.owner}"):
                st.text_input("Subject", value=item.suggested_subject, key=f"subj_{project.project_name}_{item.task_name}")
                st.text_area("Body", value=item.suggested_body, height=180, key=f"body_{project.project_name}_{item.task_name}")
                st.caption("Copy this draft into your mail client. This tool never sends email automatically.")
