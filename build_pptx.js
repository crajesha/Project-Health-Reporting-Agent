// build_pptx.js
// Generates the 7-slide monthly executive presentation from portfolio_data.json
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const portfolio = JSON.parse(fs.readFileSync(process.argv[2] || "outputs/portfolio_data.json", "utf8"));
const outPath = process.argv[3] || "outputs/presentation/Executive_Project_Health_Review.pptx";

// ---- Palette: "Midnight Executive" ----
const NAVY = "1E2761";
const ICE = "CADCFC";
const WHITE = "FFFFFF";
const RED = "C0392B";
const AMBER = "C98A1D";
const GREEN = "1E7A46";
const INK = "1F2328";
const MUTE = "5B6472";
const CARD = "F5F7FC";

const RAG_COLOR = { Green: GREEN, Amber: AMBER, Red: RED };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5
pres.author = "Project Health Reporting Agent";
pres.title = "Executive Project Health Review";

const PW = 13.3, PH = 7.5;

function addFooter(slide, pageNum) {
  slide.addText(`Professional Services  |  Confidential`, {
    x: 0.5, y: PH - 0.42, w: 6, h: 0.3, fontSize: 9, color: MUTE, fontFace: "Calibri",
  });
  slide.addText(`${pageNum}`, {
    x: PW - 1, y: PH - 0.42, w: 0.5, h: 0.3, fontSize: 9, color: MUTE, align: "right", fontFace: "Calibri",
  });
}

function titleBar(slide, kicker, title) {
  slide.background = { color: WHITE };
  slide.addText(kicker.toUpperCase(), {
    x: 0.6, y: 0.35, w: 10, h: 0.35, fontSize: 12, color: NAVY, bold: true, charSpacing: 2, fontFace: "Calibri",
  });
  slide.addText(title, {
    x: 0.6, y: 0.65, w: 11.5, h: 0.8, fontSize: 30, bold: true, color: INK, fontFace: "Cambria",
  });
}

// ============================================================
// SLIDE 1 — Title
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 9.6, y: -2.3, w: 6.5, h: 6.5, fill: { color: "273580" }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 10.8, y: 3.4, w: 4.0, h: 4.0, fill: { color: "2B3A8C" }, line: { type: "none" } });

  s.addText("EXECUTIVE PROJECT HEALTH REVIEW", {
    x: 0.7, y: 2.55, w: 10.5, h: 0.5, fontSize: 16, color: ICE, bold: true, charSpacing: 3, fontFace: "Calibri",
  });
  s.addText("Portfolio Health, Trends & Risk Outlook", {
    x: 0.7, y: 3.0, w: 10.8, h: 1.1, fontSize: 40, bold: true, color: WHITE, fontFace: "Cambria",
  });
  const monthLabel = new Date().toLocaleString("en-US", { month: "long", year: "numeric" });
  s.addText(`Monthly Synthesis — ${monthLabel}`, {
    x: 0.7, y: 4.05, w: 8, h: 0.5, fontSize: 16, color: ICE, fontFace: "Calibri",
  });
  s.addText("Prepared by the Project Health Reporting Agent  |  Professional Services", {
    x: 0.7, y: 6.7, w: 9, h: 0.4, fontSize: 11, color: "9FB1E8", fontFace: "Calibri",
  });
}

// ============================================================
// SLIDE 2 — Executive Summary
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 2", "Executive Summary");

  const total = portfolio.length;
  const redCount = portfolio.filter(p => p.rag === "Red").length;
  const amberCount = portfolio.filter(p => p.rag === "Amber").length;
  const greenCount = portfolio.filter(p => p.rag === "Green").length;
  const avgScore = (portfolio.reduce((a, p) => a + p.score, 0) / total).toFixed(0);
  const disagreements = portfolio.filter(p => !p.agrees).length;

  const stats = [
    { label: "Projects Reviewed", value: `${total}`, color: NAVY },
    { label: "Red / Amber", value: `${redCount} / ${amberCount}`, color: RED },
    { label: "Avg. Health Score", value: `${avgScore}`, color: NAVY },
    { label: "Status Overrides vs. Self-Report", value: `${disagreements}/${total}`, color: AMBER },
  ];
  const cardW = 2.85, gap = 0.25, startX = 0.6, y = 1.75;
  stats.forEach((st, i) => {
    const x = startX + i * (cardW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: 1.55, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.08 },
    });
    s.addText(st.value, { x, y: y + 0.15, w: cardW, h: 0.75, fontSize: 34, bold: true, color: st.color, align: "center", fontFace: "Cambria" });
    s.addText(st.label, { x: x + 0.1, y: y + 0.98, w: cardW - 0.2, h: 0.5, fontSize: 11, color: MUTE, align: "center", fontFace: "Calibri" });
  });

  const worst = [...portfolio].sort((a, b) => a.score - b.score)[0];
  const summaryText = [
    { text: `The portfolio's average composite health score is ${avgScore}/100. `, options: {} },
    { text: `${disagreements} of ${total} projects carry an overall status that differs from what their own status sheet self-reports`, options: { bold: true } },
    { text: ` once schedule slippage, critical-path risk, and stakeholder comments are weighed together rather than read in isolation. `, options: {} },
    { text: `${worst.name} is the portfolio's most urgent item, rated ${worst.rag} at a ${worst.score}/100 composite score.`, options: { bold: true } },
  ];
  s.addText(summaryText, { x: 0.6, y: 3.7, w: 12.1, h: 1.3, fontSize: 15, color: INK, fontFace: "Calibri", lineSpacingMultiple: 1.25 });

  s.addText("KEY TAKEAWAY FOR THE CLIENT", { x: 0.6, y: 5.25, w: 6, h: 0.3, fontSize: 11, bold: true, color: NAVY, charSpacing: 1.5, fontFace: "Calibri" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 5.6, w: 12.1, h: 1.25, rectRadius: 0.08, fill: { color: NAVY }, line: { type: "none" },
  });
  s.addText("Self-reported schedule health is no longer a reliable single indicator across this portfolio — composite, multi-signal scoring is now the basis for governance decisions and client updates.", {
    x: 0.9, y: 5.72, w: 11.5, h: 1.0, fontSize: 14, color: WHITE, italic: true, fontFace: "Calibri", valign: "middle",
  });

  addFooter(s, 2);
}

// ============================================================
// SLIDE 3 — Portfolio Health (RAG distribution + per-project bars)
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 3", "Portfolio Health at a Glance");

  // Left: donut-style RAG distribution via pie chart
  const redCount = portfolio.filter(p => p.rag === "Red").length;
  const amberCount = portfolio.filter(p => p.rag === "Amber").length;
  const greenCount = portfolio.filter(p => p.rag === "Green").length;

  s.addText("RAG DISTRIBUTION", { x: 0.6, y: 1.7, w: 4, h: 0.3, fontSize: 12, bold: true, color: NAVY, charSpacing: 1, fontFace: "Calibri" });
  s.addChart(pres.charts.DOUGHNUT, [
    { name: "RAG", labels: ["Red", "Amber", "Green"], values: [redCount, amberCount, greenCount] },
  ], {
    x: 0.5, y: 2.05, w: 4.6, h: 4.4,
    chartColors: [RED, AMBER, GREEN],
    showLegend: true, legendPos: "b", legendColor: INK, legendFontSize: 11,
    dataLabelColor: WHITE, showValue: true, dataLabelFontSize: 12, holeSize: 55,
    chartColorsOpacity: 95,
  });

  // Right: per-project composite score bar chart
  s.addText("COMPOSITE HEALTH SCORE BY PROJECT", { x: 5.5, y: 1.7, w: 7, h: 0.3, fontSize: 12, bold: true, color: NAVY, charSpacing: 1, fontFace: "Calibri" });
  s.addChart(pres.charts.BAR, [
    { name: "Score", labels: portfolio.map(p => p.name.length > 22 ? p.name.slice(0, 20) + "…" : p.name), values: portfolio.map(p => p.score) },
  ], {
    x: 5.4, y: 2.05, w: 7.3, h: 4.4,
    barDir: "bar",
    chartColors: portfolio.map(p => RAG_COLOR[p.rag]),
    valAxisMinVal: 0, valAxisMaxVal: 100,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 11, dataLabelColor: INK,
    catAxisLabelColor: INK, valAxisLabelColor: MUTE, catAxisLabelFontSize: 11,
    showLegend: false,
  });

  addFooter(s, 3);
}

// ============================================================
// SLIDE 4 — Project Trend Analysis (cross-project patterns, not per-project recap)
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 4", "Cross-Portfolio Trend Analysis");

  const avgSlip = (portfolio.reduce((a, p) => a + p.avg_delay_days, 0) / portfolio.length).toFixed(1);
  const totalOverdue = portfolio.reduce((a, p) => a + p.missed_end_dates, 0);
  const totalCriticalRisk = portfolio.reduce((a, p) => a + p.critical_red + p.critical_delayed, 0);
  const disagreeCount = portfolio.filter(p => !p.agrees).length;

  const trends = [
    {
      h: "Self-reported status is systematically unreliable",
      b: `${disagreeCount} of ${portfolio.length} projects show a computed RAG that diverges from their own Schedule Health field — in both directions (one overly optimistic, one overly pessimistic). This is a portfolio-wide pattern, not an isolated data-entry issue, and points to inconsistent local definitions of "Green" across PMs.`,
    },
    {
      h: "Critical-path risk is concentrated, not evenly spread",
      b: `${totalCriticalRisk} critical-path tasks across the portfolio are either Red or trending behind schedule. Where slippage exists, it clusters on the critical path rather than being spread evenly across the plan — the projects most at risk are at risk because of a small number of high-leverage tasks.`,
    },
    {
      h: "Slippage severity is rising with project maturity",
      b: `Average delay where slippage exists is ${avgSlip} days across the portfolio, with ${totalOverdue} open tasks already past their planned end date. Later-stage projects (higher % complete) show materially larger average delays than earlier-stage ones — schedule risk compounds rather than resolves as projects mature.`,
    },
  ];

  let y = 1.75;
  trends.forEach((t, i) => {
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.02, w: 0.45, h: 0.45, fill: { color: NAVY }, line: { type: "none" } });
    s.addText(`${i + 1}`, { x: 0.6, y: y + 0.02, w: 0.45, h: 0.45, fontSize: 18, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: "Cambria" });
    s.addText(t.h, { x: 1.25, y: y, w: 11.3, h: 0.4, fontSize: 16, bold: true, color: INK, fontFace: "Calibri" });
    s.addText(t.b, { x: 1.25, y: y + 0.42, w: 11.3, h: 0.85, fontSize: 12.5, color: MUTE, fontFace: "Calibri", lineSpacingMultiple: 1.2 });
    y += 1.6;
  });

  addFooter(s, 4);
}

// ============================================================
// SLIDE 5 — Common Risks
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 5", "Emerging & Common Risks");

  const risks = [
    { t: "Critical-path slippage", d: "Multiple projects carry delayed or Red critical-path tasks that directly threaten committed end dates.", sev: "High" },
    { t: "Pending approvals as silent blockers", d: "Stakeholder comments repeatedly reference approvals and dependencies awaiting sign-off — a leading indicator that precedes visible schedule slippage.", sev: "Medium" },
    { t: "Inconsistent self-reported RAG", d: "Local Schedule Health values are not comparable across projects, undermining portfolio-level roll-ups unless re-scored centrally.", sev: "Medium" },
    { t: "Completion-vs-timeline gap widening", d: "At least one project shows completion trailing elapsed timeline by a double-digit percentage, with no compensating acceleration visible.", sev: "High" },
  ];
  const sevColor = { High: RED, Medium: AMBER, Low: GREEN };

  const colW = 5.85, rowH = 2.25, gapX = 0.35, gapY = 0.25, startX = 0.6, startY = 1.75;
  risks.forEach((r, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = startX + col * (colW + gapX);
    const y = startY + row * (rowH + gapY);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: colW, h: rowH, rectRadius: 0.08, fill: { color: CARD }, line: { type: "none" },
      shadow: { type: "outer", color: "000000", blur: 8, offset: 2, angle: 90, opacity: 0.08 },
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.3, y: y + 0.28, w: 1.1, h: 0.34, rectRadius: 0.17, fill: { color: sevColor[r.sev] }, line: { type: "none" },
    });
    s.addText(r.sev.toUpperCase(), { x: x + 0.3, y: y + 0.28, w: 1.1, h: 0.34, fontSize: 10, bold: true, color: WHITE, align: "center", valign: "middle", fontFace: "Calibri" });
    s.addText(r.t, { x: x + 0.3, y: y + 0.75, w: colW - 0.6, h: 0.45, fontSize: 15, bold: true, color: INK, fontFace: "Calibri" });
    s.addText(r.d, { x: x + 0.3, y: y + 1.22, w: colW - 0.6, h: 0.9, fontSize: 11.5, color: MUTE, fontFace: "Calibri", lineSpacingMultiple: 1.2 });
  });

  addFooter(s, 5);
}

// ============================================================
// SLIDE 6 — Recommendations
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 6", "Recommendations");

  const recs = [
    { t: "Standardize RAG scoring centrally", d: "Adopt the composite scoring model portfolio-wide so status is comparable across PMs and projects, replacing locally-defined Schedule Health fields as the reporting source of truth." },
    { t: "Institute a critical-path escalation trigger", d: "Any critical-path task slipping more than 5 business days should auto-generate a follow-up action, rather than waiting for the next weekly cycle." },
    { t: "Close the approval-latency gap", d: "Track pending-approval comments as a distinct leading indicator and assign an SLA (e.g. 3 business days) before they convert into schedule slippage." },
    { t: "Weekly steering review for Red/Amber projects", d: "Formalize a standing weekly checkpoint for any project below an 80 composite score until it returns to Green." },
  ];

  let y = 1.8;
  recs.forEach((r, i) => {
    s.addText(`0${i + 1}`, { x: 0.6, y: y, w: 0.9, h: 1.0, fontSize: 30, bold: true, color: ICE_TEXT_COLOR(), fontFace: "Cambria" });
    s.addText(r.t, { x: 1.6, y: y + 0.02, w: 10.9, h: 0.4, fontSize: 16, bold: true, color: NAVY, fontFace: "Calibri" });
    s.addText(r.d, { x: 1.6, y: y + 0.44, w: 10.9, h: 0.55, fontSize: 12.5, color: MUTE, fontFace: "Calibri", lineSpacingMultiple: 1.2 });
    y += 1.28;
  });

  function ICE_TEXT_COLOR() { return "AEB9E0"; }

  addFooter(s, 6);
}

// ============================================================
// SLIDE 7 — Projects Requiring Attention (appendix-style detail table)
// ============================================================
{
  const s = pres.addSlide();
  titleBar(s, "Slide 7", "Projects Requiring Attention — Detail");

  const header = ["Project", "PM", "Status", "Score", "Open Overdue Tasks", "Critical Risk", "Self-Report vs. Computed"];
  const rows = portfolio
    .slice()
    .sort((a, b) => a.score - b.score)
    .map(p => [
      p.name,
      p.pm || "—",
      p.rag,
      `${p.score}`,
      `${p.missed_end_dates}`,
      `${p.critical_red + p.critical_delayed}/${p.critical_count}`,
      p.agrees ? "Matches" : `${p.sheet_status} → ${p.rag}`,
    ]);

  const tableRows = [header.map(h => ({ text: h, options: { bold: true, color: WHITE, fill: { color: NAVY }, fontSize: 11, align: "left", valign: "middle" } }))];
  rows.forEach(r => {
    tableRows.push(r.map((cell, ci) => {
      let color = INK;
      if (ci === 2) color = RAG_COLOR[cell] || INK;
      return { text: cell, options: { color, fontSize: 11, align: "left", valign: "middle", bold: ci === 2, fill: { color: WHITE } } };
    }));
  });

  s.addTable(tableRows, {
    x: 0.6, y: 1.8, w: 12.1, h: 2.6,
    colW: [3.0, 1.9, 1.1, 1.0, 1.8, 1.5, 1.8],
    border: { type: "solid", color: "E3E7F0", pt: 1 },
    autoPage: false,
  });

  s.addText("METHODOLOGY NOTE", { x: 0.6, y: 4.75, w: 5, h: 0.3, fontSize: 11, bold: true, color: NAVY, charSpacing: 1, fontFace: "Calibri" });
  s.addText(
    "RAG status is calculated from a weighted composite of schedule slippage (30%), completion vs. elapsed timeline (20%), critical-path task health (20%), stakeholder blocker signals (15%), and milestone health (15%), with override rules that cap the rating when critical-path or blocker risk is severe. Full methodology available on request.",
    { x: 0.6, y: 5.1, w: 12.1, h: 0.9, fontSize: 11.5, color: MUTE, italic: true, fontFace: "Calibri", lineSpacingMultiple: 1.25 }
  );

  s.addText(`Data as of the most recent weekly extract for each project. Portfolio of ${portfolio.length} project(s) reviewed this cycle.`, {
    x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 10, color: MUTE, fontFace: "Calibri",
  });

  addFooter(s, 7);
}

pres.writeFile({ fileName: outPath }).then(() => console.log("Written:", outPath));
