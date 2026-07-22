/* ============================================================================
   Intelligence_2.2_GovernanceScore.gs — GovernX Intelligence Platform
   PHASE 2 · UNIT 2.2 — Governance Score Engine

   The headline numbers. One scorecard row per company:
     8 domain scores (0-100) + Composite + Failure_Probability + Recovery_Probability

   Tab created:
     • Governance_Scorecard

   FUNCTIONS (the three required by Objective 5 are pure — no API cost):
     • scoreGovernance(id)             — Claude fills the 8 domain scores
     • calculateGovernanceScore(id)    — weighted mean of the 8 domains → Composite
     • calculateFailureProbability(id) — Framework(inverse)+DNA-risk+failure-severity
     • calculateRecoveryProbability(id)— Governance+Transparency+Current_Status
     • computeCompanyScores(id)        — runs the 3 calculators together
     • computeAllCompanyScores()       — same, for every company (batch, cheap)

   READS: Framework_Assessment (2.1), Company_DNA (1.4), Company_Failure_Map (1.3),
          Company_Master.Current_Status (1.1). Degrades gracefully if some are absent.

   SAFETY: additive, own constants, intelSS_(), no config.gs edits.
   ============================================================================ */


// ── Tab name ─────────────────────────────────────────────────────────────────
const SHEET_SCORECARD = { MAIN: "Governance_Scorecard" };

// ── Column map (1-based) ──────────────────────────────────────────────────────
const COL_SCORECARD = {
  COMPANY_ID           : 1,
  LEADERSHIP           : 2,
  GOVERNANCE           : 3,
  RISK                 : 4,
  COMPLIANCE           : 5,
  INNOVATION           : 6,
  PROCESSES            : 7,
  CULTURE              : 8,
  TRANSPARENCY         : 9,
  COMPOSITE            : 10,
  FAILURE_PROBABILITY  : 11,
  RECOVERY_PROBABILITY : 12
};

// ── Weights for the governance composite (tunable) ───────────────────────────
const SCORE_WEIGHTS = {
  LEADERSHIP  : 0.15, GOVERNANCE : 0.20, RISK       : 0.15, COMPLIANCE  : 0.15,
  INNOVATION  : 0.07, PROCESSES  : 0.08, CULTURE    : 0.10, TRANSPARENCY: 0.10
};

// Current_Status → recovery-favourability factor (0-100)
const STATUS_FACTOR = {
  "Active": 90, "Restructured": 70, "Acquired": 60, "Bankrupt": 25, "Defunct": 10
};

const INTEL_HEADER_BG_SC = "#1a1a2e";
const INTEL_HEADER_FG_SC = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupGovernanceScorecardTab() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  let sc = ss.getSheetByName(SHEET_SCORECARD.MAIN);
  const isNew = !sc;
  if (isNew) sc = ss.insertSheet(SHEET_SCORECARD.MAIN);

  sc.getRange(1, 1, 1, 12).setValues([[
    "Company_ID", "Leadership", "Governance", "Risk", "Compliance",
    "Innovation", "Processes", "Culture", "Transparency",
    "Composite", "Failure_Probability", "Recovery_Probability"
  ]]).setBackground(INTEL_HEADER_BG_SC).setFontColor(INTEL_HEADER_FG_SC).setFontWeight("bold");
  sc.setFrozenRows(1);
  [160, 100, 105, 80, 100, 100, 100, 90, 110, 100, 150, 165].forEach((w, i) => sc.setColumnWidth(i + 1, w));

  // Green-good scale on the 8 domains + composite (higher = better)
  addGradient_(sc, COL_SCORECARD.LEADERSHIP, 9, "#F8696B", "#FFEB84", "#63BE7B");   // 0..100 red→green
  // Failure probability: higher = worse → reverse (green→red)
  addGradient_(sc, COL_SCORECARD.FAILURE_PROBABILITY, 1, "#63BE7B", "#FFEB84", "#F8696B");
  // Recovery probability: higher = better → green
  addGradient_(sc, COL_SCORECARD.RECOVERY_PROBABILITY, 1, "#F8696B", "#FFEB84", "#63BE7B");

  ui.alert("✅ Governance Scorecard Ready",
    (isNew ? "Created Governance_Scorecard.\n\n" : "Governance_Scorecard refreshed.\n\n") +
    "Flow per company:\n" +
    "1. Score Governance Domains (Claude fills the 8 domains)\n" +
    "2. Compute Scores (Composite + Failure% + Recovery%)\n\n" +
    "Failure% and Recovery% also use Framework (2.1), DNA (1.4), Failure Map (1.3).",
    ui.ButtonSet.OK);
}

function addGradient_(sheet, startCol, numCols, minC, midC, maxC) {
  const range = sheet.getRange(2, startCol, 999, numCols);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(minC, SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMidpointWithValue(midC, SpreadsheetApp.InterpolationType.NUMBER, "50")
    .setGradientMaxpointWithValue(maxC, SpreadsheetApp.InterpolationType.NUMBER, "100")
    .setRanges([range]).build();
  const rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}


// ══════════════════════════════════════════════════════════════════════════════
// SCORE GOVERNANCE DOMAINS — Claude fills the 8 domain scores (0-100)
// ══════════════════════════════════════════════════════════════════════════════
function scoreGovernance(companyId) {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }
  const name = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();

  let context = "";
  const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
  if (profile) {
    const pRow = findProfileRow_(profile, companyId);
    if (pRow !== -1) context = "Root cause: " + profile.getRange(pRow, COL_COMPANY_PROFILE.ROOT_CAUSE).getValue();
  }

  const prompt = `
Score this company on 8 governance domains, each 0-100 (100 = best-in-class,
0 = total failure). These are QUALITY scores.

Company: ${name}
${context}

Return EXACTLY these 8 lines, digits only:
LEADERSHIP: [0-100]
GOVERNANCE: [0-100]
RISK: [0-100]
COMPLIANCE: [0-100]
INNOVATION: [0-100]
PROCESSES: [0-100]
CULTURE: [0-100]
TRANSPARENCY: [0-100]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 800);
    const num = (f) => {
      const m = raw.match(new RegExp(f + ":\\s*(\\d+)"));
      return m ? Math.max(0, Math.min(100, parseInt(m[1], 10))) : "";
    };

    const sc  = ss.getSheetByName(SHEET_SCORECARD.MAIN);
    const row = ensureScorecardRow_(sc, companyId);
    sc.getRange(row, COL_SCORECARD.LEADERSHIP  ).setValue(num("LEADERSHIP"));
    sc.getRange(row, COL_SCORECARD.GOVERNANCE  ).setValue(num("GOVERNANCE"));
    sc.getRange(row, COL_SCORECARD.RISK        ).setValue(num("RISK"));
    sc.getRange(row, COL_SCORECARD.COMPLIANCE  ).setValue(num("COMPLIANCE"));
    sc.getRange(row, COL_SCORECARD.INNOVATION  ).setValue(num("INNOVATION"));
    sc.getRange(row, COL_SCORECARD.PROCESSES   ).setValue(num("PROCESSES"));
    sc.getRange(row, COL_SCORECARD.CULTURE     ).setValue(num("CULTURE"));
    sc.getRange(row, COL_SCORECARD.TRANSPARENCY).setValue(num("TRANSPARENCY"));

    // Immediately compute the composite from the fresh domains
    calculateGovernanceScore(companyId);
    ui.alert("✅ Governance domains scored for " + name + ".\nComposite computed. " +
             "Run 'Compute Scores' to add Failure% and Recovery%.");
  } catch (err) {
    if (typeof logError === "function") logError("Intel 2.2 — Score Governance", companyId, "API/Runtime", err.message);
    ui.alert("❌ Governance scoring failed: " + err.message);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// THE THREE ENGINES (pure — no API). Each defaults to the selected company,
// writes its result to the scorecard, and returns the value.
// ══════════════════════════════════════════════════════════════════════════════

// Weighted mean of the 8 domain scores → Composite (0-100)
function calculateGovernanceScore(companyId) {
  const ss = intelSS_();
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return null;

  const sc  = ss.getSheetByName(SHEET_SCORECARD.MAIN);
  if (!sc) return null;
  const row = ensureScorecardRow_(sc, companyId);

  const domains = [
    ["LEADERSHIP", SCORE_WEIGHTS.LEADERSHIP], ["GOVERNANCE", SCORE_WEIGHTS.GOVERNANCE],
    ["RISK", SCORE_WEIGHTS.RISK], ["COMPLIANCE", SCORE_WEIGHTS.COMPLIANCE],
    ["INNOVATION", SCORE_WEIGHTS.INNOVATION], ["PROCESSES", SCORE_WEIGHTS.PROCESSES],
    ["CULTURE", SCORE_WEIGHTS.CULTURE], ["TRANSPARENCY", SCORE_WEIGHTS.TRANSPARENCY]
  ];
  let wSum = 0, vSum = 0;
  domains.forEach(([key, w]) => {
    const v = sc.getRange(row, COL_SCORECARD[key]).getValue();
    if (v !== "" && !isNaN(v)) { vSum += Number(v) * w; wSum += w; }
  });
  if (wSum === 0) return null;   // no domains scored yet

  const composite = Math.round(vSum / wSum);
  sc.getRange(row, COL_SCORECARD.COMPOSITE).setValue(composite);
  return composite;
}

// Failure probability (0-100): weak framework + risky DNA + severe failures
function calculateFailureProbability(companyId) {
  const ss = intelSS_();
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return null;

  const components = [];
  const fw = readFrameworkComposite_(ss, companyId);
  if (fw !== null) components.push({ v: 100 - fw, w: 0.5 });
  const dnaRisk = readDnaRiskIndex_(ss, companyId);
  if (dnaRisk !== null) components.push({ v: dnaRisk, w: 0.3 });
  const failSev = readFailureSeverityIndex_(ss, companyId);
  if (failSev !== null) components.push({ v: failSev, w: 0.2 });

  const fp = combineWeighted_(components);
  if (fp === null) return null;

  const sc  = ss.getSheetByName(SHEET_SCORECARD.MAIN);
  const row = ensureScorecardRow_(sc, companyId);
  sc.getRange(row, COL_SCORECARD.FAILURE_PROBABILITY).setValue(fp);
  return fp;
}

// Recovery probability (0-100): governance strength + transparency + status
function calculateRecoveryProbability(companyId) {
  const ss = intelSS_();
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return null;

  const sc  = ss.getSheetByName(SHEET_SCORECARD.MAIN);
  const row = ensureScorecardRow_(sc, companyId);

  const components = [];
  let gov = sc.getRange(row, COL_SCORECARD.COMPOSITE).getValue();
  if (gov === "" || isNaN(gov)) gov = calculateGovernanceScore(companyId);
  if (gov !== null && gov !== "" && !isNaN(gov)) components.push({ v: Number(gov), w: 0.4 });

  const trans = readDnaTrait_(ss, companyId, "transparency");
  if (trans !== null) components.push({ v: trans * 10, w: 0.3 });

  const statusFactor = readStatusFactor_(ss, companyId);
  if (statusFactor !== null) components.push({ v: statusFactor, w: 0.3 });

  const rp = combineWeighted_(components);
  if (rp === null) return null;
  sc.getRange(row, COL_SCORECARD.RECOVERY_PROBABILITY).setValue(rp);
  return rp;
}

// Run all three for one company (menu entry) + a summary alert
function computeCompanyScores(companyId) {
  const ui = SpreadsheetApp.getUi();
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const g = calculateGovernanceScore(companyId);
  const f = calculateFailureProbability(companyId);
  const r = calculateRecoveryProbability(companyId);

  ui.alert("✅ Scores computed for " + companyId,
    "Governance composite: " + (g === null ? "— (score domains first)" : g) + "\n" +
    "Failure probability:  " + (f === null ? "— (needs Framework/DNA/Failures)" : f + "%") + "\n" +
    "Recovery probability: " + (r === null ? "—" : r + "%"),
    ui.ButtonSet.OK);
}

// Batch across every company (cheap — pure calculations only)
function computeAllCompanyScores() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();
  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  if (!master) { ui.alert("Company_Master not found."); return; }

  const ids = master.getRange(2, COL_COMPANY_MASTER.COMPANY_ID, Math.max(master.getLastRow() - 1, 0), 1)
                    .getValues().map(r => (r[0] || "").toString().trim()).filter(Boolean);
  let n = 0;
  ids.forEach(id => {
    const g = calculateGovernanceScore(id);
    const f = calculateFailureProbability(id);
    const r = calculateRecoveryProbability(id);
    if (g !== null || f !== null || r !== null) n++;
  });
  ui.alert("✅ Batch complete", "Computed scores for " + n + " of " + ids.length + " companies.\n" +
    "(Companies with no domain/framework/DNA data yet were skipped.)", ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
function ensureScorecardRow_(sc, companyId) {
  const data = sc.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_SCORECARD.COMPANY_ID - 1] || "").toString().trim() === companyId) return i + 1;
  }
  const row = sc.getLastRow() + 1;
  sc.getRange(row, COL_SCORECARD.COMPANY_ID).setValue(companyId);
  return row;
}

// Weighted mean over present components, renormalized; rounded 0-100
function combineWeighted_(components) {
  if (!components.length) return null;
  let wSum = 0, vSum = 0;
  components.forEach(c => { vSum += c.v * c.w; wSum += c.w; });
  if (wSum === 0) return null;
  return Math.max(0, Math.min(100, Math.round(vSum / wSum)));
}

// Mean of the 7 Framework_Assessment stage scores (0-100), or null
function readFrameworkComposite_(ss, companyId) {
  const a = ss.getSheetByName("Framework_Assessment");
  if (!a) return null;
  const data = a.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || "").toString().trim() === companyId) {
      const vals = data[i].slice(1, 8).map(Number).filter(v => !isNaN(v));
      if (!vals.length) return null;
      return vals.reduce((s, v) => s + v, 0) / vals.length;
    }
  }
  return null;
}

// DNA risk index (0-100): mean(RiskAppetite, 11-BoardIndep, 11-ComplianceMat, 11-Transparency)*10
function readDnaRiskIndex_(ss, companyId) {
  const d = readDnaRow_(ss, companyId);
  if (!d) return null;
  const parts = [];
  if (d.risk_appetite       != null) parts.push(d.risk_appetite);
  if (d.board_independence  != null) parts.push(11 - d.board_independence);
  if (d.compliance_maturity != null) parts.push(11 - d.compliance_maturity);
  if (d.transparency        != null) parts.push(11 - d.transparency);
  if (!parts.length) return null;
  return (parts.reduce((s, v) => s + v, 0) / parts.length) * 10;
}

function readDnaTrait_(ss, companyId, trait) {
  const d = readDnaRow_(ss, companyId);
  return d && d[trait] != null ? d[trait] : null;
}

function readDnaRow_(ss, companyId) {
  const dna = ss.getSheetByName("Company_DNA");
  if (!dna || typeof COL_DNA === "undefined") return null;
  const data = dna.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_DNA.COMPANY_ID - 1] || "").toString().trim() === companyId) {
      const g = (col) => { const v = Number(data[i][col - 1]); return isNaN(v) ? null : v; };
      return {
        decision_speed        : g(COL_DNA.DECISION_SPEED),
        risk_appetite         : g(COL_DNA.RISK_APPETITE),
        innovation            : g(COL_DNA.INNOVATION),
        centralization        : g(COL_DNA.CENTRALIZATION),
        board_independence    : g(COL_DNA.BOARD_INDEPENDENCE),
        compliance_maturity   : g(COL_DNA.COMPLIANCE_MATURITY),
        operational_complexity: g(COL_DNA.OPERATIONAL_COMPLEXITY),
        transparency          : g(COL_DNA.TRANSPARENCY)
      };
    }
  }
  return null;
}

// Avg severity of mapped failures × 10 → 0-100, or null
function readFailureSeverityIndex_(ss, companyId) {
  const m = ss.getSheetByName("Company_Failure_Map");
  if (!m || typeof COL_FAILURE_MAP === "undefined") return null;
  const data = m.getDataRange().getValues();
  const sevs = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_FAILURE_MAP.COMPANY_ID - 1] || "").toString().trim() === companyId) {
      const s = Number(data[i][COL_FAILURE_MAP.SEVERITY - 1]);
      if (!isNaN(s)) sevs.push(s);
    }
  }
  if (!sevs.length) return null;
  return Math.min(100, (sevs.reduce((a, b) => a + b, 0) / sevs.length) * 10);
}

// Current_Status → factor (0-100), default 50 if status blank/unknown
function readStatusFactor_(ss, companyId) {
  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const row = findCompanyRow_(master, companyId);
  if (row === -1) return null;
  const status = (master.getRange(row, COL_COMPANY_MASTER.CURRENT_STATUS).getValue() || "").toString().trim();
  if (!status) return 50;
  return STATUS_FACTOR[status] != null ? STATUS_FACTOR[status] : 50;
}
