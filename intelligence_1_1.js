/* ============================================================================
   Intelligence_1.1_CompanyEntity.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.1 — Company Entity

   Creates the entity data layer: every company becomes an object.
   Tabs created:
     • Company_Master        (the object — one row per company)
     • Company_Profile       (the analysis — one row per company)
     • Company_Timeline      (dated events — many rows per company)
     • Company_Relationships (company-to-company links — many rows)

   SAFETY:
   - Purely additive. Touches NO existing tab, column, or function.
   - Defines its OWN constants (SHEET_COMPANY, COL_COMPANY_*) — does not
     modify config.gs, so there is no redeclaration risk.
   - Does NOT define onOpen() (that would clash with menu.gs). Use
     buildIntelligenceMenu() or paste the menu lines shown at the bottom.

   HOW TO USE (first time):
   1. Paste this file into the Apps Script editor as a new file.
   2. Run  setupCompanyEntityTabs()   → creates the 4 tabs.
   3. Run  backfillCompaniesFromDB()  → seeds companies from DB_S + DB_F.
   4. (Optional, per company) run enrichCompanyProfile("CO-0001")
      and buildCompanyTimeline("CO-0001") to fill the analysis + timeline.
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_COMPANY = {
  MASTER        : "Company_Master",
  PROFILE       : "Company_Profile",
  TIMELINE      : "Company_Timeline",
  RELATIONSHIPS : "Company_Relationships"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_COMPANY_MASTER = {
  COMPANY_ID      : 1,
  COMPANY_NAME    : 2,
  COUNTRY         : 3,
  INDUSTRY        : 4,
  DOMAIN          : 5,
  FOUNDED         : 6,
  COLLAPSED       : 7,
  CEO             : 8,
  BOARD_CHAIR     : 9,
  EMPLOYEES       : 10,
  PEAK_REVENUE    : 11,
  PEAK_MARKET_CAP : 12,
  CURRENT_STATUS  : 13,
  WIKIPEDIA       : 14,
  WEBSITE         : 15,
  CREATED_AT      : 16,
  UPDATED_AT      : 17
};

const COL_COMPANY_PROFILE = {
  COMPANY_ID          : 1,
  SUMMARY             : 2,
  COLLAPSE_TYPE       : 3,
  ROOT_CAUSE          : 4,
  GOVERNANCE_FAILURES : 5,
  RISK_FAILURES       : 6,
  COMPLIANCE_FAILURES : 7,
  BPR_ISSUES          : 8,
  LESSONS_LEARNED     : 9,
  VIDEO_COUNT         : 10
};

const COL_COMPANY_TIMELINE = {
  COMPANY_ID : 1,
  DATE       : 2,
  EVENT      : 3,
  CATEGORY   : 4,
  IMPACT     : 5,
  SOURCE     : 6
};

const COL_COMPANY_REL = {
  COMPANY_A    : 1,
  RELATIONSHIP : 2,
  COMPANY_B    : 3,
  STRENGTH     : 4
};

// ── Controlled vocabularies (used for dropdown validation) ───────────────────
const COMPANY_STATUS_ENUM   = ["Active", "Bankrupt", "Acquired", "Restructured", "Defunct"];
const COLLAPSE_TYPE_ENUM     = ["Slow Decline", "Sudden Collapse", "Scandal", "Acquisition", "Turnaround"];
const TIMELINE_CATEGORY_ENUM = ["Founding", "Growth", "Warning", "Decision", "Shock", "Collapse", "Aftermath"];
const TIMELINE_IMPACT_ENUM   = ["Low", "Medium", "High", "Critical"];
const RELATIONSHIP_ENUM      = ["similar_to", "competitor_of", "acquired_by", "spun_off", "same_pattern_as", "regulated_by_same"];

// ── Header styling (matches the GovernX dark-header convention) ───────────────
const INTEL_HEADER_BG = "#1a1a2e";
const INTEL_HEADER_FG = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP — create the 4 tabs (idempotent; never wipes existing data)
// ══════════════════════════════════════════════════════════════════════════════
function setupCompanyEntityTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const created = [];

  created.push(ensureTab_(ss, SHEET_COMPANY.MASTER, [
    "Company_ID", "Company_Name", "Country", "Industry", "Domain",
    "Founded", "Collapsed", "CEO", "Board_Chair", "Employees",
    "Peak_Revenue", "Peak_Market_Cap", "Current_Status",
    "Wikipedia", "Website", "Created_At", "Updated_At"
  ], [160, 200, 120, 150, 90, 80, 80, 160, 160, 100, 130, 140, 130, 220, 220, 150, 150]));

  created.push(ensureTab_(ss, SHEET_COMPANY.PROFILE, [
    "Company_ID", "Summary", "Collapse_Type", "Root_Cause",
    "Governance_Failures", "Risk_Failures", "Compliance_Failures",
    "BPR_Issues", "Lessons_Learned", "Video_Count"
  ], [160, 380, 150, 300, 300, 300, 300, 300, 300, 100]));

  created.push(ensureTab_(ss, SHEET_COMPANY.TIMELINE, [
    "Company_ID", "Date", "Event", "Category", "Impact", "Source"
  ], [160, 110, 420, 130, 110, 260]));

  created.push(ensureTab_(ss, SHEET_COMPANY.RELATIONSHIPS, [
    "Company_A", "Relationship", "Company_B", "Strength"
  ], [160, 180, 160, 100]));

  // ── Dropdown validations on controlled-vocabulary columns ────────────────
  applyColumnEnum_(ss, SHEET_COMPANY.MASTER,   COL_COMPANY_MASTER.CURRENT_STATUS, COMPANY_STATUS_ENUM);
  applyColumnEnum_(ss, SHEET_COMPANY.PROFILE,  COL_COMPANY_PROFILE.COLLAPSE_TYPE, COLLAPSE_TYPE_ENUM);
  applyColumnEnum_(ss, SHEET_COMPANY.TIMELINE, COL_COMPANY_TIMELINE.CATEGORY,     TIMELINE_CATEGORY_ENUM);
  applyColumnEnum_(ss, SHEET_COMPANY.TIMELINE, COL_COMPANY_TIMELINE.IMPACT,       TIMELINE_IMPACT_ENUM);
  applyColumnEnum_(ss, SHEET_COMPANY.RELATIONSHIPS, COL_COMPANY_REL.RELATIONSHIP, RELATIONSHIP_ENUM);

  const madeNew = created.filter(Boolean);
  ui.alert(
    "✅ Company Entity Tabs Ready",
    (madeNew.length
      ? "Created: " + madeNew.join(", ") + "\n\n"
      : "All 4 tabs already existed — headers/validation refreshed.\n\n") +
    "Next: run backfillCompaniesFromDB() to seed companies from DB_S + DB_F.",
    ui.ButtonSet.OK
  );
}

// ── Create a tab with a styled header row if it does not already exist ────────
// Returns the tab name if newly created, or "" if it already existed.
function ensureTab_(ss, name, headers, widths) {
  let sheet = ss.getSheetByName(name);
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(name);

  // Write/refresh header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground(INTEL_HEADER_BG)
    .setFontColor(INTEL_HEADER_FG)
    .setFontWeight("bold");
  sheet.setFrozenRows(1);

  if (widths && widths.length) {
    widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return isNew ? name : "";
}

// ── Apply a dropdown (requireValueInList) to an entire data column ────────────
function applyColumnEnum_(ss, sheetName, col, list) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)   // allow blanks / manual override
    .build();
  // Rows 2..1000 for this column
  sheet.getRange(2, col, 999, 1).setDataValidation(rule);
}


// ══════════════════════════════════════════════════════════════════════════════
// ID GENERATION — next CO-#### (scans Company_Master, globally sequential)
// ══════════════════════════════════════════════════════════════════════════════
function generateCompanyId_(masterSheet) {
  const data = masterSheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][COL_COMPANY_MASTER.COMPANY_ID - 1] || "").toString().trim();
    const m  = id.match(/^CO-(\d+)$/);
    if (m) {
      const seq = parseInt(m[1], 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }
  return "CO-" + String(maxSeq + 1).padStart(4, "0");
}


// ══════════════════════════════════════════════════════════════════════════════
// BACKFILL — seed Company_Master from the existing DB_S + DB_F databases
// Header-driven (matches columns by name), so it is resilient to layout drift.
// Only maps directly-available facts (Name, Country, Domain, Industry).
// Financials/CEO/dates are left blank for later Claude enrichment — no fabrication.
// De-duplicates by normalized company name across DB_S, DB_F, and Company_Master.
// ══════════════════════════════════════════════════════════════════════════════
function backfillCompaniesFromDB() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ui     = SpreadsheetApp.getUi();
  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);

  if (!master) {
    ui.alert("Run setupCompanyEntityTabs() first — Company_Master is missing.");
    return;
  }

  // Build the set of names already in Company_Master
  const existing = new Set();
  const mData = master.getDataRange().getValues();
  for (let i = 1; i < mData.length; i++) {
    const nm = mData[i][COL_COMPANY_MASTER.COMPANY_NAME - 1];
    if (nm) existing.add(normalizeName_(nm));
  }

  const now = new Date();
  const newRows = [];

  ["DB_S", "DB_F"].forEach(dbName => {
    const db = ss.getSheetByName(dbName);
    if (!db) return;

    const data = db.getDataRange().getValues();
    if (data.length < 2) return;

    const idx = headerIndexMap_(data[0]); // name(lower) -> 0-based col index

    // Resolve source columns by best-matching header names
    const cName     = pickCol_(idx, ["company", "company name", "name"]);
    const cCountry  = pickCol_(idx, ["country"]);
    const cDomain   = pickCol_(idx, ["domain"]);
    const cIndustry = pickCol_(idx, ["industry"]);

    if (cName === -1) return; // cannot seed without a name column

    for (let i = 1; i < data.length; i++) {
      const rawName = (data[i][cName] || "").toString().trim();
      if (!rawName) continue;

      const key = normalizeName_(rawName);
      if (existing.has(key)) continue; // dedupe
      existing.add(key);

      const id = "CO-" + String(1 + master.getLastRow() - 1 + newRows.length).padStart(4, "0");

      const row = new Array(17).fill("");
      row[COL_COMPANY_MASTER.COMPANY_ID   - 1] = id;
      row[COL_COMPANY_MASTER.COMPANY_NAME - 1] = rawName;
      row[COL_COMPANY_MASTER.COUNTRY      - 1] = cCountry  !== -1 ? (data[i][cCountry]  || "") : "";
      row[COL_COMPANY_MASTER.INDUSTRY     - 1] = cIndustry !== -1 ? (data[i][cIndustry] || "") : "";
      row[COL_COMPANY_MASTER.DOMAIN       - 1] = cDomain   !== -1 ? (data[i][cDomain]   || "") : "";
      row[COL_COMPANY_MASTER.CREATED_AT   - 1] = now;
      row[COL_COMPANY_MASTER.UPDATED_AT   - 1] = now;
      newRows.push(row);
    }
  });

  if (newRows.length === 0) {
    ui.alert("Backfill complete — no new companies found (all already present, or DB_S/DB_F not found).");
    return;
  }

  // Re-sequence IDs cleanly against the true current max (guards against reruns)
  const startSeq = maxCompanySeq_(master) + 1;
  newRows.forEach((r, i) => {
    r[COL_COMPANY_MASTER.COMPANY_ID - 1] = "CO-" + String(startSeq + i).padStart(4, "0");
  });

  master.getRange(master.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);

  ui.alert(
    "✅ Backfill Complete",
    "Added " + newRows.length + " companies to Company_Master (" +
    "IDs CO-" + String(startSeq).padStart(4, "0") + " … CO-" +
    String(startSeq + newRows.length - 1).padStart(4, "0") + ").\n\n" +
    "Name, Country, Domain, and Industry were mapped from DB_S/DB_F.\n" +
    "Financials, CEO, and dates were left blank — fill via enrichment or manually.",
    ui.ButtonSet.OK
  );
}

// Highest CO-#### sequence currently in Company_Master
function maxCompanySeq_(master) {
  const data = master.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const m = (data[i][COL_COMPANY_MASTER.COMPANY_ID - 1] || "").toString().match(/^CO-(\d+)$/);
    if (m) { const s = parseInt(m[1], 10); if (s > maxSeq) maxSeq = s; }
  }
  return maxSeq;
}

// Build header-name -> 0-based column index (lowercased, trimmed)
function headerIndexMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = (h || "").toString().trim().toLowerCase();
    if (key && !(key in map)) map[key] = i;
  });
  return map;
}

// Return the 0-based column for the first matching candidate name, else -1
function pickCol_(idx, candidates) {
  for (const c of candidates) {
    if (c in idx) return idx[c];
  }
  return -1;
}

// Normalize a company name for de-duplication
function normalizeName_(name) {
  return name.toString()
    .toLowerCase()
    .replace(/\b(inc|corp|corporation|company|co|ltd|limited|plc|llc|group|holdings|sa|ag|nv)\b/g, "")
    .replace(/[^a-z0-9]+/g, "")
    .trim();
}

// Find the Company_Master row (1-based) for a given Company_ID, or -1
function findCompanyRow_(master, companyId) {
  const data = master.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_COMPANY_MASTER.COMPANY_ID - 1] || "").toString().trim() === companyId) {
      return i + 1;
    }
  }
  return -1;
}


// ══════════════════════════════════════════════════════════════════════════════
// CLAUDE ENRICHMENT — Intelligence system context (separate persona from the
// content-engine SYSTEM_CONTEXT; this one is an analyst, not a scriptwriter)
// ══════════════════════════════════════════════════════════════════════════════
const INTEL_SYSTEM_CONTEXT = `
You are the GovernX Intelligence Analyst. You build a structured governance
intelligence database about companies — their governance, risk, compliance, and
process failures.

RULES:
- Report only what is well-established and verifiable. No speculation as fact.
- For living individuals, use cautious framing ("reported", "according to").
- No fabricated quotes or invented figures. If a fact is unknown, leave it blank.
- Be concise and factual. Output ONLY the requested fields in the exact format.
`;

// ── Enrich Company_Profile for one company ────────────────────────────────────
function enrichCompanyProfile(companyId) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ui     = SpreadsheetApp.getUi();

  // If run with no argument (e.g. the editor Run button), use the selected row
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);

  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name     = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();
  const country  = master.getRange(mRow, COL_COMPANY_MASTER.COUNTRY).getValue();
  const industry = master.getRange(mRow, COL_COMPANY_MASTER.INDUSTRY).getValue();

  const prompt = `
Build the governance intelligence profile for this company.

Company : ${name}
Country : ${country || "unknown"}
Industry: ${industry || "unknown"}

Return EXACTLY these fields, one value per line (leave blank if genuinely unknown):

SUMMARY: [3-4 sentence factual overview of the company and what happened to it]
COLLAPSE_TYPE: [one of: Slow Decline | Sudden Collapse | Scandal | Acquisition | Turnaround]
ROOT_CAUSE: [1-2 sentences — the single deepest cause]
GOVERNANCE_FAILURES: [semicolon-separated list]
RISK_FAILURES: [semicolon-separated list]
COMPLIANCE_FAILURES: [semicolon-separated list]
BPR_ISSUES: [semicolon-separated process/operating-model failures]
LESSONS_LEARNED: [semicolon-separated GRC/BPR lessons]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 3000);
    const get = (f) => {
      const m = raw.match(new RegExp(f + ":\\s*([\\s\\S]*?)(?=\\n[A-Z_]{3,}:|$)"));
      return m ? m[1].replace(/\n/g, " ").trim() : "";
    };

    const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
    let pRow = findProfileRow_(profile, companyId);
    if (pRow === -1) pRow = profile.getLastRow() + 1;

    profile.getRange(pRow, COL_COMPANY_PROFILE.COMPANY_ID         ).setValue(companyId);
    profile.getRange(pRow, COL_COMPANY_PROFILE.SUMMARY            ).setValue(get("SUMMARY"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.COLLAPSE_TYPE      ).setValue(get("COLLAPSE_TYPE"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.ROOT_CAUSE         ).setValue(get("ROOT_CAUSE"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.GOVERNANCE_FAILURES).setValue(get("GOVERNANCE_FAILURES"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.RISK_FAILURES      ).setValue(get("RISK_FAILURES"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.COMPLIANCE_FAILURES).setValue(get("COMPLIANCE_FAILURES"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.BPR_ISSUES         ).setValue(get("BPR_ISSUES"));
    profile.getRange(pRow, COL_COMPANY_PROFILE.LESSONS_LEARNED    ).setValue(get("LESSONS_LEARNED"));
    if (!profile.getRange(pRow, COL_COMPANY_PROFILE.VIDEO_COUNT).getValue()) {
      profile.getRange(pRow, COL_COMPANY_PROFILE.VIDEO_COUNT).setValue(0);
    }

    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());
    ui.alert("✅ Profile enriched for " + name + " (" + companyId + ").");

  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.1 — Enrich Profile", companyId, "API/Runtime", err.message);
    ui.alert("❌ Enrichment failed: " + err.message);
  }
}

function findProfileRow_(profile, companyId) {
  const data = profile.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_COMPANY_PROFILE.COMPANY_ID - 1] || "").toString().trim() === companyId) return i + 1;
  }
  return -1;
}

// ── Enrich the factual reference columns in Company_Master for one company ────
// Fills Founded, Collapsed, CEO, Board_Chair, Employees, Peak_Revenue,
// Peak_Market_Cap, Current_Status, Wikipedia, Website. Only writes a cell when
// Claude returns a value (unknown fields stay blank — no fabrication, and any
// manual edits you already made to a filled cell are preserved).
function enrichCompanyMaster(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // If run with no argument (e.g. the editor Run button), use the selected row
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name     = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();
  const country  = master.getRange(mRow, COL_COMPANY_MASTER.COUNTRY).getValue();
  const industry = master.getRange(mRow, COL_COMPANY_MASTER.INDUSTRY).getValue();

  const prompt = `
Fill the factual reference fields for this company. Report only well-established
facts. If a value is genuinely unknown or not applicable, leave it blank.

Company : ${name}
Country : ${country || "unknown"}
Industry: ${industry || "unknown"}

Return EXACTLY these fields, one value per line:

FOUNDED: [year founded, e.g. 1975]
COLLAPSED: [year of collapse/bankruptcy/acquisition, or blank if still operating]
CEO: [most relevant CEO — current, or the one at the key inflection point]
BOARD_CHAIR: [board chair name, or blank]
EMPLOYEES: [approx headcount at peak, digits only, e.g. 145000]
PEAK_REVENUE: [peak annual revenue, e.g. $60.9B (2023)]
PEAK_MARKET_CAP: [peak market capitalization, e.g. $3.3T (2024)]
CURRENT_STATUS: [one of: Active | Bankrupt | Acquired | Restructured | Defunct]
WIKIPEDIA: [full Wikipedia URL]
WEBSITE: [official website URL]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 1500);
    const get = (f) => { const m = raw.match(new RegExp(f + ":\\s*(.+)")); return m ? m[1].trim() : ""; };
    const setIf = (col, val) => { if (val) master.getRange(mRow, col).setValue(val); };

    setIf(COL_COMPANY_MASTER.FOUNDED,         get("FOUNDED"));
    setIf(COL_COMPANY_MASTER.COLLAPSED,       get("COLLAPSED"));
    setIf(COL_COMPANY_MASTER.CEO,             get("CEO"));
    setIf(COL_COMPANY_MASTER.BOARD_CHAIR,     get("BOARD_CHAIR"));
    setIf(COL_COMPANY_MASTER.EMPLOYEES,       get("EMPLOYEES"));
    setIf(COL_COMPANY_MASTER.PEAK_REVENUE,    get("PEAK_REVENUE"));
    setIf(COL_COMPANY_MASTER.PEAK_MARKET_CAP, get("PEAK_MARKET_CAP"));
    setIf(COL_COMPANY_MASTER.CURRENT_STATUS,  get("CURRENT_STATUS"));
    setIf(COL_COMPANY_MASTER.WIKIPEDIA,       get("WIKIPEDIA"));
    setIf(COL_COMPANY_MASTER.WEBSITE,         get("WEBSITE"));
    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());

    ui.alert("✅ Master facts filled for " + name + " (" + companyId + ").");
  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.1 — Enrich Master", companyId, "API/Runtime", err.message);
    ui.alert("❌ Enrichment failed: " + err.message);
  }
}

// ── Build Company_Timeline for one company ────────────────────────────────────
function buildCompanyTimeline(companyId) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ui     = SpreadsheetApp.getUi();

  // If run with no argument (e.g. the editor Run button), use the selected row
  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const name = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();

  const prompt = `
Build a factual event timeline for the company below — the key dated moments in
its governance/risk/collapse story, in chronological order.

Company: ${name}

Return 6-12 events, each as a block in EXACTLY this format:

EVENT_START
DATE: [YYYY or YYYY-MM-DD]
EVENT: [one factual sentence]
CATEGORY: [one of: Founding | Growth | Warning | Decision | Shock | Collapse | Aftermath]
IMPACT: [one of: Low | Medium | High | Critical]
SOURCE: [publication/report name or URL, or "General record"]
EVENT_END
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 4000);
    const blocks = raw.match(/EVENT_START([\s\S]*?)EVENT_END/g) || [];
    if (!blocks.length) { ui.alert("No timeline events returned for " + name + "."); return; }

    const timeline = ss.getSheetByName(SHEET_COMPANY.TIMELINE);

    // Remove existing rows for this company (clean rebuild)
    const tData = timeline.getDataRange().getValues();
    for (let i = tData.length - 1; i >= 1; i--) {
      if ((tData[i][COL_COMPANY_TIMELINE.COMPANY_ID - 1] || "").toString().trim() === companyId) {
        timeline.deleteRow(i + 1);
      }
    }

    const rows = blocks.map(b => {
      const g = (f) => { const m = b.match(new RegExp(f + ":\\s*(.+)")); return m ? m[1].trim() : ""; };
      return [companyId, g("DATE"), g("EVENT"), g("CATEGORY"), g("IMPACT"), g("SOURCE")];
    });

    timeline.getRange(timeline.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());
    ui.alert("✅ Timeline built for " + name + " — " + rows.length + " events.");

  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.1 — Timeline", companyId, "API/Runtime", err.message);
    ui.alert("❌ Timeline build failed: " + err.message);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// OPTIONAL MENU — run this to add an Intelligence menu for the current session,
// OR paste the .addItem lines below into your existing onOpen() in menu.gs.
// ══════════════════════════════════════════════════════════════════════════════
function buildIntelligenceMenu() {
  SpreadsheetApp.getUi()
    .createMenu("🧠 GovernX Intelligence")
    // ── Unit 1.1 — Company Entity ────────────────────────────────────────────
    .addItem("① Setup Company Entity Tabs", "setupCompanyEntityTabs")
    .addItem("② Backfill Companies (DB_S + DB_F)", "backfillCompaniesFromDB")
    .addSeparator()
    .addItem("Enrich Master Facts — selected company", "enrichSelectedCompanyMaster_")
    .addItem("Enrich Profile — selected company", "enrichSelectedCompanyProfile_")
    .addItem("Build Timeline — selected company", "buildSelectedCompanyTimeline_")
    // ── Unit 1.3 — Failure Taxonomy ──────────────────────────────────────────
    .addSeparator()
    .addItem("③ Setup Failure Taxonomy", "setupFailureTaxonomyTabs")
    .addItem("Map Failures — selected company", "mapCompanyFailures")
    // ── Unit 1.4 — Company DNA ───────────────────────────────────────────────
    .addSeparator()
    .addItem("④ Setup Company DNA", "setupCompanyDNATab")
    .addItem("Score DNA — selected company", "scoreCompanyDNA")
    // ── Unit 1.5 — Pattern Engine ────────────────────────────────────────────
    .addSeparator()
    .addItem("⑤ Setup Pattern Engine", "setupPatternEngineTabs")
    .addItem("Detect Patterns — selected company", "detectPatterns")
    .addItem("Detect Patterns — ALL companies", "detectAllPatterns")
    // ── Unit 1.2 — Knowledge Graph ───────────────────────────────────────────
    .addSeparator()
    .addItem("⑥ Setup Knowledge Graph", "setupKnowledgeGraphTabs")
    .addItem("Sync Company Nodes (all)", "syncCompanyNodes")
    .addItem("Generate Graph — selected company", "generateGraphFromCompany")
    // ── Unit 1.6 — Living Database ───────────────────────────────────────────
    .addSeparator()
    .addItem("⑦ Setup Living Database", "setupCompanyUpdatesTab")
    .addItem("Add Update — selected company", "addCompanyUpdate")
    // ── Unit 2.1 — GovernX Framework ─────────────────────────────────────────
    .addSeparator()
    .addItem("⑧ Setup GovernX Framework", "setupFrameworkTabs")
    .addItem("Assess Framework — selected company", "assessFramework")
    // ── Unit 2.2 — Governance Score Engine ───────────────────────────────────
    .addSeparator()
    .addItem("⑨ Setup Governance Scorecard", "setupGovernanceScorecardTab")
    .addItem("Score Governance Domains — selected company", "scoreGovernance")
    .addItem("Compute Scores — selected company", "computeCompanyScores")
    .addItem("Compute Scores — ALL companies", "computeAllCompanyScores")
    // ── Unit 3.1 — Mr. X Character ───────────────────────────────────────────
    .addSeparator()
    .addItem("⑩ Setup Mr. X Character", "setupMrXTabs")
    .addItem("Check Mr. X Art Ready", "checkMrXArtReady")
    .addItem("Build Character Blueprint — selected video", "buildCharacterBlueprint")
    .addItem("List HeyGen Talking Photos (get ID)", "heyGenListTalkingPhotos")
    .addItem("Stage 7C — Generate Mr. X Presenter", "generateMrXPresenter")
    // ── Unit 4.1 — Revenue Engine ────────────────────────────────────────────
    .addSeparator()
    .addItem("⑪ Setup Revenue Tabs", "setupRevenueTabs")
    .addItem("Add Digital Product", "addDigitalProduct")
    .addItem("Stage 10B — Commercial Layer (selected video)", "generateCommercialLayer")
    // ── Unit 5.1 — Audience Memory ───────────────────────────────────────────
    .addSeparator()
    .addItem("⑫ Setup Audience Memory", "setupAudienceMemoryTab")
    .addItem("Update Audience Memory (learn)", "updateAudienceMemory")
    // ── Unit 5.2 — Dashboards ────────────────────────────────────────────────
    .addSeparator()
    .addItem("📊 Build Intelligence Dashboards", "buildIntelligenceDashboards")
    // Evidence work now lives in the 🔎 Research menu (research_bridge.gs),
    // backed by the Node verified-claims engine. The old in-Apps-Script
    // Evidence Graph and the manual Pilot Ledger were removed — Apps Script
    // cannot parse PDFs or survive a long fetch, and the pilot ledger's claims
    // were hand-typed rather than document-verified.
    // ── Utilities ────────────────────────────────────────────────────────────
    .addSeparator()
    .addItem("🎨 Organize Intelligence Tabs (color + move to end)", "organizeIntelligenceTabs")
    .addItem("🙈 Hide Intelligence Tabs", "hideIntelligenceTabs")
    .addItem("👁️ Show Intelligence Tabs", "showIntelligenceTabs")
    .addToUi();
}

// Convenience wrappers: act on the Company_ID in the currently-selected row
function enrichSelectedCompanyMaster_() {
  const id = getSelectedCompanyId_();
  if (id) enrichCompanyMaster(id);
}
function enrichSelectedCompanyProfile_() {
  const id = getSelectedCompanyId_();
  if (id) enrichCompanyProfile(id);
}
function buildSelectedCompanyTimeline_() {
  const id = getSelectedCompanyId_();
  if (id) buildCompanyTimeline(id);
}
function getSelectedCompanyId_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui    = SpreadsheetApp.getUi();
  if (sheet.getName() !== SHEET_COMPANY.MASTER) {
    ui.alert("Select a row in the Company_Master tab first."); return null;
  }
  const row = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("Select a data row (not the header)."); return null; }
  const id = sheet.getRange(row, COL_COMPANY_MASTER.COMPANY_ID).getValue();
  if (!id) { ui.alert("That row has no Company_ID."); return null; }
  return id.toString().trim();
}

/* ── To make the menu permanent, add these lines inside onOpen() in menu.gs ──
   (right before .addToUi(); leaves everything else untouched):

     .addSeparator()
     .addItem("🧠 Setup Company Entity Tabs", "setupCompanyEntityTabs")
     .addItem("🧠 Backfill Companies (DB_S + DB_F)", "backfillCompaniesFromDB")

   ...or simply run buildIntelligenceMenu() once per session for a separate menu.
--------------------------------------------------------------------------- */
