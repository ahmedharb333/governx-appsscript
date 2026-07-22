/* ============================================================================
   Intelligence_1.3_FailureTaxonomy.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.3 — Failure Taxonomy

   Creates a controlled library of failure types, and a map linking each
   company to the failures it exhibits (with a severity score).

   Tabs created:
     • Failure_Taxonomy   (the library — one row per named failure, FL-###)
     • Company_Failure_Map(assignments — Company_ID × Failure_ID × Severity)

   SAFETY:
   - Purely additive. Touches NO existing tab, column, or function.
   - Own constants (SHEET_FAILURE, COL_FAILURE_*). No config.gs edits.
   - Reuses helpers already defined in Intelligence_1.1 (getSelectedCompanyId_,
     findCompanyRow_, INTEL_SYSTEM_CONTEXT, callClaudeWithCustomSystem, logError).
     → Keep the Unit 1.1 file in the project; this file depends on it.

   HOW TO USE:
   1. Run  setupFailureTaxonomyTabs()  → creates + seeds the taxonomy (52 failures).
   2. Select a company row in Company_Master, run  mapCompanyFailures()
      → Claude assigns the relevant Failure_IDs + severities for that company.
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_FAILURE = {
  TAXONOMY : "Failure_Taxonomy",
  MAP      : "Company_Failure_Map"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_FAILURE_TAXONOMY = {
  FAILURE_ID   : 1,
  FAILURE_NAME : 2,
  DESCRIPTION  : 3,
  CATEGORY     : 4
};

const COL_FAILURE_MAP = {
  COMPANY_ID : 1,
  FAILURE_ID : 2,
  SEVERITY   : 3
};

// ── The 13 failure categories (Objective 3) ──────────────────────────────────
const FAILURE_CATEGORIES = [
  "Board Failure", "Leadership Failure", "Governance Failure", "Risk Failure",
  "Compliance Failure", "Culture Failure", "Innovation Failure",
  "Technology Failure", "Ethics Failure", "Audit Failure", "M&A Failure",
  "Communication Failure", "Regulatory Failure"
];

// ── Seed library: [Failure_Name, Description, Category] ───────────────────────
// FL-### IDs are assigned automatically by row order at setup time.
const FAILURE_SEED = [
  // Board
  ["Passive Board",            "Board failed to challenge management or provide real oversight.",        "Board Failure"],
  ["Founder-Captured Board",   "Board dominated by the founder or insiders; lacked independence.",       "Board Failure"],
  ["Overboarded Directors",    "Directors too stretched across mandates to oversee effectively.",        "Board Failure"],
  ["Related-Party Board",      "Conflicts of interest among directors compromised objectivity.",         "Board Failure"],
  // Leadership
  ["Autocratic CEO",           "Unchecked CEO power with no room for dissent or challenge.",             "Leadership Failure"],
  ["Succession Void",          "No credible leadership succession plan in place.",                       "Leadership Failure"],
  ["Visionary Blind Spot",     "Leadership ignored disconfirming evidence about its strategy.",          "Leadership Failure"],
  ["Empire Building",          "Growth and ego prioritized over sound strategy.",                        "Leadership Failure"],
  // Governance
  ["Weak Oversight Structure", "No effective governance mechanisms or accountability lines.",            "Governance Failure"],
  ["Concentrated Control",     "Dual-class or ownership concentration removed accountability.",          "Governance Failure"],
  ["Missing Independent Committees", "No independent audit or risk committee.",                          "Governance Failure"],
  ["Governance-in-Name-Only",  "Policies existed on paper but were never enforced.",                     "Governance Failure"],
  // Risk
  ["No Risk Function",         "Absent or immature enterprise risk management.",                         "Risk Failure"],
  ["Ignored Warning Signals",  "Known risks were not escalated or acted upon.",                          "Risk Failure"],
  ["Concentration Risk",       "Overexposure to a single customer, product, or geography.",              "Risk Failure"],
  ["Tail-Risk Blindness",      "Low-probability, high-impact events were underestimated.",               "Risk Failure"],
  // Compliance
  ["Regulatory Breach",        "Violated applicable laws or regulations.",                               "Compliance Failure"],
  ["Weak Controls",            "Inadequate internal controls over key processes.",                       "Compliance Failure"],
  ["Sanctions/AML Failure",    "Money-laundering or sanctions-screening breakdowns.",                    "Compliance Failure"],
  ["Disclosure Failure",       "Misleading, incomplete, or omitted required disclosures.",               "Compliance Failure"],
  // Culture
  ["Toxic Performance Culture","Extreme pressure drove employees toward misconduct.",                    "Culture Failure"],
  ["Fear-Based Silence",       "Employees could not safely raise concerns.",                             "Culture Failure"],
  ["Normalization of Deviance","Small violations became accepted as routine.",                           "Culture Failure"],
  ["Ethics-Blind Incentives",  "Incentive schemes rewarded the wrong behavior.",                         "Culture Failure"],
  // Innovation
  ["Innovation Blindness",     "Failed to see or adopt disruptive technology.",                          "Innovation Failure"],
  ["Cannibalization Fear",     "Protected a legacy product instead of backing the future.",              "Innovation Failure"],
  ["Slow Adaptation",          "Too slow to pivot to a clear market shift.",                             "Innovation Failure"],
  ["R&D Misallocation",        "Invested heavily in the wrong technology bets.",                         "Innovation Failure"],
  // Technology
  ["Legacy Lock-In",           "Trapped by outdated core systems.",                                      "Technology Failure"],
  ["Platform Missed",          "Missed a platform or ecosystem shift.",                                  "Technology Failure"],
  ["Security Failure",         "Major breach, data loss, or systemic security lapse.",                   "Technology Failure"],
  ["Tech Debt Collapse",       "Systems could not scale or adapt when it mattered.",                     "Technology Failure"],
  // Ethics
  ["Fraud",                    "Deliberate financial misrepresentation.",                                "Ethics Failure"],
  ["Deception of Stakeholders","Misled customers, investors, or the public.",                            "Ethics Failure"],
  ["Conflict of Interest",     "Self-dealing or undisclosed personal interest.",                         "Ethics Failure"],
  ["Cover-Up",                 "Concealed known problems from stakeholders.",                             "Ethics Failure"],
  // Audit
  ["Auditor Capture",          "External auditor was not genuinely independent.",                        "Audit Failure"],
  ["Missed Red Flags",         "Audit process failed to detect material issues.",                        "Audit Failure"],
  ["Internal Audit Weakness",  "No effective internal audit function.",                                  "Audit Failure"],
  ["Accounting Manipulation",  "Earnings management or off-balance-sheet engineering.",                  "Audit Failure"],
  // M&A
  ["Overpaid Acquisition",     "Destroyed value by overpaying for a deal.",                              "M&A Failure"],
  ["Integration Failure",      "Failed post-merger integration.",                                        "M&A Failure"],
  ["Due Diligence Gap",        "Missed material risks before closing a deal.",                            "M&A Failure"],
  ["Culture Clash",            "Incompatible cultures wrecked a merger.",                                "M&A Failure"],
  // Communication
  ["Crisis Mismanagement",     "Mishandled communications during a crisis.",                             "Communication Failure"],
  ["Stakeholder Misalignment", "Failed to align or inform key stakeholders.",                            "Communication Failure"],
  ["Transparency Deficit",     "Withheld material information from stakeholders.",                        "Communication Failure"],
  ["Reputation Mismanagement", "Failed to protect stakeholder trust.",                                   "Communication Failure"],
  // Regulatory
  ["Regulatory Capture",       "Improperly influenced its own regulator.",                               "Regulatory Failure"],
  ["Missed Regulatory Shift",  "Failed to adapt to new regulation.",                                      "Regulatory Failure"],
  ["Licensing/Approval Failure","Operated without proper approvals or licenses.",                        "Regulatory Failure"],
  ["Cross-Border Compliance Gap","Failed to meet multi-jurisdiction requirements.",                      "Regulatory Failure"]
];

const INTEL_HEADER_BG_FL = "#1a1a2e";
const INTEL_HEADER_FG_FL = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP — create the 2 tabs and seed the taxonomy (seed runs only if empty)
// ══════════════════════════════════════════════════════════════════════════════
function setupFailureTaxonomyTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── Failure_Taxonomy ──────────────────────────────────────────────────────
  let tax = ss.getSheetByName(SHEET_FAILURE.TAXONOMY);
  const taxIsNew = !tax;
  if (taxIsNew) tax = ss.insertSheet(SHEET_FAILURE.TAXONOMY);

  tax.getRange(1, 1, 1, 4).setValues([["Failure_ID", "Failure_Name", "Description", "Category"]])
     .setBackground(INTEL_HEADER_BG_FL).setFontColor(INTEL_HEADER_FG_FL).setFontWeight("bold");
  tax.setFrozenRows(1);
  [130, 240, 460, 190].forEach((w, i) => tax.setColumnWidth(i + 1, w));

  // Seed only when there is no data yet (avoids duplicate FL rows on rerun)
  let seeded = 0;
  if (tax.getLastRow() < 2) {
    const rows = FAILURE_SEED.map((f, i) => [
      "FL-" + String(i + 1).padStart(3, "0"), f[0], f[1], f[2]
    ]);
    tax.getRange(2, 1, rows.length, 4).setValues(rows);
    seeded = rows.length;
  }

  // Category dropdown on the Category column
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(FAILURE_CATEGORIES, true).setAllowInvalid(true).build();
  tax.getRange(2, COL_FAILURE_TAXONOMY.CATEGORY, Math.max(tax.getLastRow() - 1, 999), 1)
     .setDataValidation(catRule);

  // ── Company_Failure_Map ───────────────────────────────────────────────────
  let map = ss.getSheetByName(SHEET_FAILURE.MAP);
  const mapIsNew = !map;
  if (mapIsNew) map = ss.insertSheet(SHEET_FAILURE.MAP);

  map.getRange(1, 1, 1, 3).setValues([["Company_ID", "Failure_ID", "Severity"]])
     .setBackground(INTEL_HEADER_BG_FL).setFontColor(INTEL_HEADER_FG_FL).setFontWeight("bold");
  map.setFrozenRows(1);
  [160, 130, 100].forEach((w, i) => map.setColumnWidth(i + 1, w));

  ui.alert(
    "✅ Failure Taxonomy Ready",
    (taxIsNew ? "Created Failure_Taxonomy. " : "Failure_Taxonomy refreshed. ") +
    (mapIsNew ? "Created Company_Failure_Map.\n\n" : "Company_Failure_Map refreshed.\n\n") +
    (seeded ? "Seeded " + seeded + " failure types across " + FAILURE_CATEGORIES.length + " categories.\n\n"
            : "Taxonomy already had data — seed skipped.\n\n") +
    "Next: select a company in Company_Master and run mapCompanyFailures().",
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// MAP — Claude assigns the applicable Failure_IDs + severity for one company
// Pulls the company's known failures from Company_Profile (if present) to inform
// the mapping, and constrains Claude to the existing taxonomy IDs only.
// ══════════════════════════════════════════════════════════════════════════════
function mapCompanyFailures(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();  // from Intelligence_1.1
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);    // from Intelligence_1.1
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();

  // Load the taxonomy so Claude can only choose existing IDs
  const tax = ss.getSheetByName(SHEET_FAILURE.TAXONOMY);
  if (!tax || tax.getLastRow() < 2) {
    ui.alert("Run setupFailureTaxonomyTabs() first — the taxonomy is empty.");
    return;
  }
  const taxData  = tax.getRange(2, 1, tax.getLastRow() - 1, 4).getValues();
  const validIds = new Set(taxData.map(r => (r[0] || "").toString().trim()));
  const taxList  = taxData.map(r => `${r[0]} | ${r[1]} | ${r[3]}`).join("\n");

  // Pull known failures from Company_Profile (optional context)
  let profileContext = "";
  const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
  if (profile) {
    const pRow = findProfileRow_(profile, companyId);   // from Intelligence_1.1
    if (pRow !== -1) {
      const g = (col) => profile.getRange(pRow, col).getValue();
      profileContext =
        "Known analysis for context:\n" +
        "Root cause: " + g(COL_COMPANY_PROFILE.ROOT_CAUSE) + "\n" +
        "Governance failures: " + g(COL_COMPANY_PROFILE.GOVERNANCE_FAILURES) + "\n" +
        "Risk failures: " + g(COL_COMPANY_PROFILE.RISK_FAILURES) + "\n" +
        "Compliance failures: " + g(COL_COMPANY_PROFILE.COMPLIANCE_FAILURES) + "\n" +
        "BPR issues: " + g(COL_COMPANY_PROFILE.BPR_ISSUES) + "\n";
    }
  }

  const prompt = `
Map the company below to the failure types it genuinely exhibited.

Company: ${name}
${profileContext}

Choose ONLY from this taxonomy (use the exact Failure_ID):
${taxList}

For each applicable failure, output a block. Assign Severity 1-10
(10 = central/defining cause, 1 = minor/peripheral). Select the 4-10 most
relevant — do not force-fit failures that do not clearly apply.

MAP_START
FAILURE_ID: [FL-###]
SEVERITY: [1-10]
MAP_END
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 2000);
    const blocks = raw.match(/MAP_START([\s\S]*?)MAP_END/g) || [];

    const rows = [];
    blocks.forEach(b => {
      const idM  = b.match(/FAILURE_ID:\s*(FL-\d+)/i);
      const sevM = b.match(/SEVERITY:\s*(\d+)/i);
      if (!idM) return;
      const id = idM[1].toUpperCase();
      if (!validIds.has(id)) return;                 // reject IDs not in taxonomy
      let sev = sevM ? parseInt(sevM[1], 10) : 5;
      sev = Math.max(1, Math.min(10, sev));
      rows.push([companyId, id, sev]);
    });

    if (!rows.length) { ui.alert("No valid failures returned for " + name + "."); return; }

    // Clean rebuild for this company
    const map = ss.getSheetByName(SHEET_FAILURE.MAP);
    const mData = map.getDataRange().getValues();
    for (let i = mData.length - 1; i >= 1; i--) {
      if ((mData[i][COL_FAILURE_MAP.COMPANY_ID - 1] || "").toString().trim() === companyId) {
        map.deleteRow(i + 1);
      }
    }
    map.getRange(map.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());

    ui.alert("✅ Mapped " + rows.length + " failures for " + name + " (" + companyId + ").");

  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.3 — Map Failures", companyId, "API/Runtime", err.message);
    ui.alert("❌ Failure mapping failed: " + err.message);
  }
}
