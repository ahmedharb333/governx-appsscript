/* ============================================================================
   Intelligence_2.1_Framework.gs — GovernX Intelligence Platform
   PHASE 2 · UNIT 2.1 — GovernX Framework Engine

   The GovernX causal model: a 7-stage chain where each stage enables the next.
     Leadership → Culture → Governance → Risk → Processes → Execution → Outcome
   Weak early stages propagate downstream — this is the spine of the whole thesis.

   Tabs created:
     • GovernX_Framework   (the 7-stage reference model — seeded)
     • Framework_Assessment(per-company score 0-100 for each of the 7 stages)

   SAFETY:
   - Additive. Own constants (SHEET_FRAMEWORK, COL_FRAMEWORK_*). No config.gs edits.
   - Uses intelSS_() and reuses Intelligence_1.1 helpers + optional context from
     Company_Profile / Company_DNA / Company_Failure_Map. Logs each assessment to
     Company_Updates (1.6) when that tab exists.

   HOW TO USE:
   1. Run  setupFrameworkTabs()   → creates + seeds the 7-stage model.
   2. Select a company, run  assessFramework()  → Claude scores the 7 stages.
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_FRAMEWORK = {
  MODEL      : "GovernX_Framework",
  ASSESSMENT : "Framework_Assessment"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_FRAMEWORK_MODEL = {
  STAGE_ORDER        : 1,
  STAGE_NAME         : 2,
  DEFINITION         : 3,
  LEADING_INDICATORS : 4
};

const COL_FRAMEWORK_ASSESS = {
  COMPANY_ID       : 1,
  LEADERSHIP_SCORE : 2,
  CULTURE_SCORE    : 3,
  GOVERNANCE_SCORE : 4,
  RISK_SCORE       : 5,
  PROCESS_SCORE    : 6,
  EXECUTION_SCORE  : 7,
  OUTCOME_SCORE    : 8
};

// ── The 7-stage reference model (seeded once) ────────────────────────────────
// [order, name, definition, leading indicators]
const FRAMEWORK_SEED = [
  [1, "Leadership", "The quality, integrity, and accountability of those at the top.",
      "CEO power balance; succession depth; tone from the top; board challenge"],
  [2, "Culture", "The shared norms and incentives that actually shape behavior.",
      "Incentive alignment; psychological safety; whistleblower activity; ethics"],
  [3, "Governance", "The oversight structures that hold power accountable.",
      "Board independence; committee structure; ownership concentration; disclosure"],
  [4, "Risk", "How the organization identifies, escalates, and manages risk.",
      "Risk function maturity; escalation paths; risk appetite discipline"],
  [5, "Processes", "The operating model and controls that turn strategy into action.",
      "Control effectiveness; standardization; operational complexity; automation"],
  [6, "Execution", "How reliably the organization delivers on its decisions.",
      "Delivery consistency; adaptability; operational KPIs; responsiveness"],
  [7, "Outcome", "The realized result: financial, reputational, and survival.",
      "Financial health; market position; stakeholder trust; longevity"]
];

const INTEL_HEADER_BG_FW = "#1a1a2e";
const INTEL_HEADER_FG_FW = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupFrameworkTabs() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  // ── GovernX_Framework (reference model) ───────────────────────────────────
  let model = ss.getSheetByName(SHEET_FRAMEWORK.MODEL);
  const modelNew = !model;
  if (modelNew) model = ss.insertSheet(SHEET_FRAMEWORK.MODEL);
  model.getRange(1, 1, 1, 4).setValues([["Stage_Order", "Stage_Name", "Definition", "Leading_Indicators"]])
       .setBackground(INTEL_HEADER_BG_FW).setFontColor(INTEL_HEADER_FG_FW).setFontWeight("bold");
  model.setFrozenRows(1);
  [110, 140, 420, 420].forEach((w, i) => model.setColumnWidth(i + 1, w));

  let seeded = 0;
  if (model.getLastRow() < 2) {
    model.getRange(2, 1, FRAMEWORK_SEED.length, 4).setValues(FRAMEWORK_SEED);
    seeded = FRAMEWORK_SEED.length;
  }

  // ── Framework_Assessment (per-company scores) ─────────────────────────────
  let assess = ss.getSheetByName(SHEET_FRAMEWORK.ASSESSMENT);
  const assessNew = !assess;
  if (assessNew) assess = ss.insertSheet(SHEET_FRAMEWORK.ASSESSMENT);
  assess.getRange(1, 1, 1, 8).setValues([[
    "Company_ID", "Leadership_Score", "Culture_Score", "Governance_Score",
    "Risk_Score", "Process_Score", "Execution_Score", "Outcome_Score"
  ]]).setBackground(INTEL_HEADER_BG_FW).setFontColor(INTEL_HEADER_FG_FW).setFontWeight("bold");
  assess.setFrozenRows(1);
  [160, 140, 130, 150, 110, 130, 140, 130].forEach((w, i) => assess.setColumnWidth(i + 1, w));
  applyScoreColorScale_(assess, 2, 7);   // color scale on the 7 score columns

  ui.alert("✅ GovernX Framework Ready",
    (seeded ? "Seeded the 7-stage model.\n" : "Model already present.\n") +
    (assessNew ? "Created Framework_Assessment.\n\n" : "Framework_Assessment refreshed.\n\n") +
    "Next: select a company and run assessFramework().",
    ui.ButtonSet.OK);
}

// Red→amber→green gradient 0-100 over a block of score columns
function applyScoreColorScale_(sheet, startCol, numCols) {
  const range = sheet.getRange(2, startCol, 999, numCols);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#F8696B", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMidpointWithValue("#FFEB84", SpreadsheetApp.InterpolationType.NUMBER, "50")
    .setGradientMaxpointWithValue("#63BE7B", SpreadsheetApp.InterpolationType.NUMBER, "100")
    .setRanges([range]).build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}


// ══════════════════════════════════════════════════════════════════════════════
// ASSESS — Claude scores the 7 framework stages (0-100) for one company
// ══════════════════════════════════════════════════════════════════════════════
function assessFramework(companyId) {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();

  // Optional grounding context from prior units
  let context = "";
  const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
  if (profile) {
    const pRow = findProfileRow_(profile, companyId);
    if (pRow !== -1) {
      context += "Summary: " + profile.getRange(pRow, COL_COMPANY_PROFILE.SUMMARY).getValue() + "\n";
      context += "Root cause: " + profile.getRange(pRow, COL_COMPANY_PROFILE.ROOT_CAUSE).getValue() + "\n";
    }
  }

  const prompt = `
Assess this company against the GovernX 7-stage framework. Each stage enables the
next: Leadership → Culture → Governance → Risk → Processes → Execution → Outcome.

Score EACH stage 0-100, where 100 = exemplary strength at that stage and 0 = total
failure. For a collapsed company, later stages (especially Outcome) will be low;
be honest about where in the chain the weakness began.

Company: ${name}
${context}

Return EXACTLY these 7 lines, digits only:
LEADERSHIP: [0-100]
CULTURE: [0-100]
GOVERNANCE: [0-100]
RISK: [0-100]
PROCESSES: [0-100]
EXECUTION: [0-100]
OUTCOME: [0-100]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 800);
    const num = (f) => {
      const m = raw.match(new RegExp(f + ":\\s*(\\d+)"));
      if (!m) return "";
      return Math.max(0, Math.min(100, parseInt(m[1], 10)));
    };

    const scores = {
      L: num("LEADERSHIP"), C: num("CULTURE"), G: num("GOVERNANCE"),
      R: num("RISK"), P: num("PROCESSES"), E: num("EXECUTION"), O: num("OUTCOME")
    };

    const assess = ss.getSheetByName(SHEET_FRAMEWORK.ASSESSMENT);
    let row = findFrameworkRow_(assess, companyId);
    if (row === -1) row = assess.getLastRow() + 1;

    assess.getRange(row, COL_FRAMEWORK_ASSESS.COMPANY_ID      ).setValue(companyId);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.LEADERSHIP_SCORE).setValue(scores.L);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.CULTURE_SCORE   ).setValue(scores.C);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.GOVERNANCE_SCORE).setValue(scores.G);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.RISK_SCORE      ).setValue(scores.R);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.PROCESS_SCORE   ).setValue(scores.P);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.EXECUTION_SCORE ).setValue(scores.E);
    assess.getRange(row, COL_FRAMEWORK_ASSESS.OUTCOME_SCORE   ).setValue(scores.O);

    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());

    // Audit trail in the Living Database (if it exists)
    if (typeof logCompanyUpdate === "function" && ss.getSheetByName("Company_Updates")) {
      try {
        logCompanyUpdate(companyId, "Data",
          "Framework assessed — L:" + scores.L + " C:" + scores.C + " G:" + scores.G +
          " R:" + scores.R + " P:" + scores.P + " E:" + scores.E + " O:" + scores.O,
          "GovernX Framework Engine");
      } catch (e) { /* non-fatal */ }
    }

    ui.alert("✅ Framework assessed for " + name,
      "L:" + scores.L + "  C:" + scores.C + "  G:" + scores.G + "  R:" + scores.R +
      "  P:" + scores.P + "  E:" + scores.E + "  O:" + scores.O, ui.ButtonSet.OK);

  } catch (err) {
    if (typeof logError === "function") logError("Intel 2.1 — Framework", companyId, "API/Runtime", err.message);
    ui.alert("❌ Framework assessment failed: " + err.message);
  }
}

function findFrameworkRow_(assess, companyId) {
  const data = assess.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_FRAMEWORK_ASSESS.COMPANY_ID - 1] || "").toString().trim() === companyId) return i + 1;
  }
  return -1;
}
