/* ============================================================================
   Intelligence_5.2_Dashboards.gs — GovernX Intelligence Platform
   PHASE 5 · UNIT 5.2 — Intelligence Dashboards (Objective 13)

   Builds ONE live-formula "Intelligence_Dashboard" tab with 8 sections:
     1. Governance Index        5. Audience Intelligence
     2. Failure Patterns        6. Knowledge Graph Metrics
     3. Revenue Analytics       7. Mr. X Performance
     4. Affiliate Revenue       8. Product Revenue

   All panels are QUERY/COUNTIF/SUM formulas over the intelligence tabs, so they
   auto-update as data changes — no manual refresh. Every formula is IFERROR-
   wrapped so missing/empty tabs show a friendly note instead of an error.

   SAFETY: additive, own constants, intelSS_(), no config.gs edits. Separate from
   the production buildDashboard() — that dashboard is untouched.

   NOTE: QUERY joins work because everything is in one spreadsheet. If you later
   split intelligence into its own file, these become native again there (or use
   IMPORTRANGE). See INTEL_SPREADSHEET_ID in Intelligence_Core.

   HOW TO USE:  run  buildIntelligenceDashboards()
   ============================================================================ */


const SHEET_INTEL_DASH = "Intelligence_Dashboard";

const DASH_HEADER_BG  = "#1E293B";
const DASH_SECTION_BG = "#0F1B2E";
const DASH_ACCENT     = "#CC0000";
const DASH_MUTED      = "#6B7280";


function buildIntelligenceDashboards() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  let d = ss.getSheetByName(SHEET_INTEL_DASH);
  if (d) { d.clear(); d.clearConditionalFormatRules(); }
  else   { d = ss.insertSheet(SHEET_INTEL_DASH); }

  d.setHiddenGridlines(true);
  d.setColumnWidth(1, 20);
  for (let c = 2; c <= 6; c++) d.setColumnWidth(c, 170);

  let row = 2;

  // ── Banner ────────────────────────────────────────────────────────────────
  d.getRange(row, 2, 1, 5).merge()
    .setValue("🧠  GovernX Intelligence — Dashboards")
    .setFontSize(16).setFontWeight("bold").setFontColor("#FFFFFF")
    .setBackground(DASH_HEADER_BG).setVerticalAlignment("middle");
  d.setRowHeight(row, 44);
  row += 2;

  // ── Section helpers ───────────────────────────────────────────────────────
  const section = (title) => {
    d.getRange(row, 2, 1, 5).merge()
      .setValue(title.toUpperCase())
      .setFontSize(10).setFontWeight("bold").setFontColor(DASH_ACCENT)
      .setBackground(DASH_SECTION_BG);
    d.setRowHeight(row, 26);
    row += 1;
  };
  const headers = (labels) => {
    labels.forEach((l, i) => d.getRange(row, 2 + i).setValue(l)
      .setFontWeight("bold").setFontColor(DASH_MUTED).setFontSize(9));
    row += 1;
  };
  const formula = (f, reserve) => {
    d.getRange(row, 2).setFormula(f).setFontSize(10);
    row += (reserve || 12);
  };
  const kv = (label, f) => {
    d.getRange(row, 2, 1, 2).merge().setValue(label).setFontColor(DASH_MUTED).setFontSize(10);
    d.getRange(row, 4).setFormula(f).setFontWeight("bold").setFontSize(12).setFontColor(DASH_ACCENT);
    row += 1;
  };

  // ── 1. Governance Index (highest failure risk first) ──────────────────────
  section("1 · Governance Index — highest failure risk");
  headers(["Company_ID", "Composite", "Failure %", "Recovery %"]);
  formula('=IFERROR(QUERY(Governance_Scorecard!A2:L,"select A, J, K, L where K is not null order by K desc limit 12",0),"Run Compute Scores (Unit 2.2) first")', 13);

  // ── 2. Failure Patterns (companies per pattern) ───────────────────────────
  section("2 · Failure Patterns — companies per pattern");
  headers(["Pattern_ID", "Companies"]);
  formula('=IFERROR(QUERY(Pattern_Company_Map!A2:C,"select A, count(C) group by A order by count(C) desc label count(C) \'\'",0),"Run Detect Patterns (Unit 1.5) first")', 7);

  // ── 3. Revenue Analytics (totals) ─────────────────────────────────────────
  section("3 · Revenue Analytics");
  kv("Total affiliate revenue", '=IFERROR("$"&TEXT(SUM(Affiliate_Assets!H2:H),"#,##0"),"$0")');
  kv("Videos monetized",        '=IFERROR(COUNTA(Affiliate_Assets!A2:A),0)');
  kv("Digital products",        '=IFERROR(COUNTA(Digital_Products!A2:A),0)');
  row += 1;

  // ── 4. Affiliate Revenue (per video) ──────────────────────────────────────
  section("4 · Affiliate Revenue — by video");
  headers(["Content_ID", "Revenue"]);
  formula('=IFERROR(QUERY(Affiliate_Assets!A2:H,"select A, H where H is not null and H>0 order by H desc limit 10",0),"No revenue logged yet")', 11);

  // ── 5. Audience Intelligence ──────────────────────────────────────────────
  section("5 · Audience Intelligence — what performs best");
  headers(["Metric", "Value", "Confidence", "Sample"]);
  formula('=IFERROR(QUERY(Audience_Memory!A2:D,"select A, B, C, D where B is not null",0),"Run Update Audience Memory (Unit 5.1) first")', 8);

  // ── 6. Knowledge Graph Metrics ────────────────────────────────────────────
  section("6 · Knowledge Graph");
  kv("Total nodes", '=IFERROR(COUNTA(Knowledge_Graph!A2:A),0)');
  kv("Total edges", '=IFERROR(COUNTA(Graph_Edges!A2:A),0)');
  row += 1;
  headers(["Most-connected node", "Edges"]);
  formula('=IFERROR(QUERY(Graph_Edges!A2:D,"select A, count(C) group by A order by count(C) desc limit 8",0),"Generate Graph (Unit 1.2) first")', 9);

  // ── 7. Mr. X Performance (expression usage across directed scenes) ────────
  section("7 · Mr. X Performance — expression usage");
  headers(["Expression", "Scenes"]);
  formula('=IFERROR(QUERY(Character_Blueprint!A2:B,"select B, count(A) group by B order by count(A) desc",0),"Build Character Blueprint (Unit 3.2) first")', 8);

  // ── 8. Product Revenue ────────────────────────────────────────────────────
  section("8 · Product Revenue — by type");
  headers(["Type", "Count", "Total Price"]);
  formula('=IFERROR(QUERY(Digital_Products!A2:D,"select B, count(A), sum(C) group by B",0),"Add Digital Products (Unit 4.1) first")', 8);

  // ── Footer ────────────────────────────────────────────────────────────────
  d.getRange(row, 2, 1, 5).merge()
    .setValue("Live formulas — auto-updates as intelligence data changes.")
    .setFontSize(9).setFontColor(DASH_MUTED).setBackground(DASH_SECTION_BG);

  // Put it near the front and color its tab
  ss.setActiveSheet(d);
  d.setTabColor(typeof INTEL_TAB_COLOR !== "undefined" ? INTEL_TAB_COLOR : "#1F3A5F");

  SpreadsheetApp.flush();
  ui.alert("✅ Intelligence Dashboards Built",
    "Created the Intelligence_Dashboard tab with 8 live panels.\n\n" +
    "Each panel fills in as you run the matching unit — sections show a hint " +
    "until their data exists.",
    ui.ButtonSet.OK);
}
