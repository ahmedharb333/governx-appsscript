/* ============================================================================
   Intelligence_1.4_CompanyDNA.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.4 — Company DNA

   The behavioral fingerprint of a company: 8 traits, each scored 1-10.
   These are DESCRIPTIVE traits (how the org behaves) — NOT quality scores.
   They power pattern detection later (e.g. Founder Syndrome = high Risk_Appetite
   + low Board_Independence + high Centralization).

   Tab created:
     • Company_DNA (one row per company; 8 trait scores)

   SAFETY:
   - Purely additive. No existing tab/column/function is modified.
   - Own constants (SHEET_DNA, COL_DNA). No config.gs edits.
   - Reuses helpers from Intelligence_1.1 (getSelectedCompanyId_,
     findCompanyRow_, INTEL_SYSTEM_CONTEXT, callClaudeWithCustomSystem, logError)
     and reads Company_Profile if present for sharper scoring.

   HOW TO USE:
   1. Run  setupCompanyDNATab()                → creates the tab.
   2. Select a company in Company_Master, run  scoreCompanyDNA()
      → Claude scores the 8 traits and writes/updates that company's row.
   ============================================================================ */


// ── Tab name ─────────────────────────────────────────────────────────────────
const SHEET_DNA = { MAIN: "Company_DNA" };

// ── Column map (1-based) ──────────────────────────────────────────────────────
const COL_DNA = {
  COMPANY_ID             : 1,
  DECISION_SPEED         : 2,
  RISK_APPETITE          : 3,
  INNOVATION             : 4,
  CENTRALIZATION         : 5,
  BOARD_INDEPENDENCE     : 6,
  COMPLIANCE_MATURITY    : 7,
  OPERATIONAL_COMPLEXITY : 8,
  TRANSPARENCY           : 9
};

const INTEL_HEADER_BG_DNA = "#1a1a2e";
const INTEL_HEADER_FG_DNA = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP — create the Company_DNA tab (idempotent; never wipes data)
// ══════════════════════════════════════════════════════════════════════════════
function setupCompanyDNATab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let dna = ss.getSheetByName(SHEET_DNA.MAIN);
  const isNew = !dna;
  if (isNew) dna = ss.insertSheet(SHEET_DNA.MAIN);

  dna.getRange(1, 1, 1, 9).setValues([[
    "Company_ID", "Decision_Speed", "Risk_Appetite", "Innovation",
    "Centralization", "Board_Independence", "Compliance_Maturity",
    "Operational_Complexity", "Transparency"
  ]]).setBackground(INTEL_HEADER_BG_DNA).setFontColor(INTEL_HEADER_FG_DNA).setFontWeight("bold");
  dna.setFrozenRows(1);
  [160, 130, 120, 110, 130, 160, 160, 180, 130].forEach((w, i) => dna.setColumnWidth(i + 1, w));

  // Optional: light conditional color scale 1-10 on the 8 trait columns
  const traitRange = dna.getRange(2, 2, 999, 8);
  const rule = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue("#F8696B", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFEB84", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setGradientMaxpointWithValue("#63BE7B", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setRanges([traitRange])
    .build();
  const rules = dna.getConditionalFormatRules();
  rules.push(rule);
  dna.setConditionalFormatRules(rules);

  ui.alert(
    "✅ Company DNA Ready",
    (isNew ? "Created Company_DNA.\n\n" : "Company_DNA refreshed.\n\n") +
    "Traits are behavioral (1-10), not good/bad:\n" +
    "• Decision_Speed 1=slow → 10=fast\n" +
    "• Risk_Appetite 1=conservative → 10=aggressive\n" +
    "• Centralization 1=decentralized → 10=top-down\n" +
    "• Board_Independence 1=captured → 10=fully independent\n" +
    "• Compliance_Maturity 1=weak → 10=robust\n" +
    "• Transparency 1=opaque → 10=transparent\n\n" +
    "Next: select a company in Company_Master and run scoreCompanyDNA().",
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// SCORE — Claude scores the 8 DNA traits for one company (1-10 each)
// ══════════════════════════════════════════════════════════════════════════════
function scoreCompanyDNA(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();   // from Intelligence_1.1
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);     // from Intelligence_1.1
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name     = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();
  const industry = master.getRange(mRow, COL_COMPANY_MASTER.INDUSTRY).getValue();

  // Pull profile context if available
  let profileContext = "";
  const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
  if (profile) {
    const pRow = findProfileRow_(profile, companyId);    // from Intelligence_1.1
    if (pRow !== -1) {
      profileContext = "Analysis context: " +
        profile.getRange(pRow, COL_COMPANY_PROFILE.SUMMARY).getValue() + " " +
        profile.getRange(pRow, COL_COMPANY_PROFILE.ROOT_CAUSE).getValue();
    }
  }

  const prompt = `
Score this company's behavioral DNA. Each trait is a NEUTRAL descriptor of how
the organization behaves — not a good/bad rating. Score each 1-10.

Company : ${name}
Industry: ${industry || "unknown"}
${profileContext}

Scales:
DECISION_SPEED: 1 = slow/bureaucratic ... 10 = fast/agile
RISK_APPETITE: 1 = very conservative ... 10 = very aggressive
INNOVATION: 1 = stagnant ... 10 = highly innovative
CENTRALIZATION: 1 = highly decentralized ... 10 = highly top-down
BOARD_INDEPENDENCE: 1 = captured/no independence ... 10 = fully independent
COMPLIANCE_MATURITY: 1 = weak/immature ... 10 = robust/mature
OPERATIONAL_COMPLEXITY: 1 = simple ... 10 = extremely complex
TRANSPARENCY: 1 = opaque ... 10 = highly transparent

Return EXACTLY these 8 lines, digits only:
DECISION_SPEED: [1-10]
RISK_APPETITE: [1-10]
INNOVATION: [1-10]
CENTRALIZATION: [1-10]
BOARD_INDEPENDENCE: [1-10]
COMPLIANCE_MATURITY: [1-10]
OPERATIONAL_COMPLEXITY: [1-10]
TRANSPARENCY: [1-10]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 800);
    const num = (f) => {
      const m = raw.match(new RegExp(f + ":\\s*(\\d+)"));
      if (!m) return "";
      return Math.max(1, Math.min(10, parseInt(m[1], 10)));
    };

    const dna = ss.getSheetByName(SHEET_DNA.MAIN);
    let row = findDnaRow_(dna, companyId);
    if (row === -1) row = dna.getLastRow() + 1;

    dna.getRange(row, COL_DNA.COMPANY_ID            ).setValue(companyId);
    dna.getRange(row, COL_DNA.DECISION_SPEED        ).setValue(num("DECISION_SPEED"));
    dna.getRange(row, COL_DNA.RISK_APPETITE         ).setValue(num("RISK_APPETITE"));
    dna.getRange(row, COL_DNA.INNOVATION            ).setValue(num("INNOVATION"));
    dna.getRange(row, COL_DNA.CENTRALIZATION        ).setValue(num("CENTRALIZATION"));
    dna.getRange(row, COL_DNA.BOARD_INDEPENDENCE    ).setValue(num("BOARD_INDEPENDENCE"));
    dna.getRange(row, COL_DNA.COMPLIANCE_MATURITY   ).setValue(num("COMPLIANCE_MATURITY"));
    dna.getRange(row, COL_DNA.OPERATIONAL_COMPLEXITY).setValue(num("OPERATIONAL_COMPLEXITY"));
    dna.getRange(row, COL_DNA.TRANSPARENCY          ).setValue(num("TRANSPARENCY"));

    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());
    ui.alert("✅ DNA scored for " + name + " (" + companyId + ").");

  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.4 — Score DNA", companyId, "API/Runtime", err.message);
    ui.alert("❌ DNA scoring failed: " + err.message);
  }
}

function findDnaRow_(dna, companyId) {
  const data = dna.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_DNA.COMPANY_ID - 1] || "").toString().trim() === companyId) return i + 1;
  }
  return -1;
}
