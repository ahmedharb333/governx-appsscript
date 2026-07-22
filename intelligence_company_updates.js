/* ============================================================================
   Intelligence_1.6_CompanyUpdates.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.6 — Living Database

   An append-only change log so every company stays updateable over time
   (news, legal developments, financial results, leadership changes, data
   corrections). This is what makes the intelligence database "living".

   Tab created:
     • Company_Updates (Company_ID, Date, Update_Type, Description, Source)

   SAFETY:
   - Additive. Own constants (SHEET_UPDATES, COL_UPDATES). No config.gs edits.
   - Uses intelSS_() (Intelligence_Core) and reuses helpers from Intelligence_1.1
     (getSelectedCompanyId_, findCompanyRow_). Keep those files in the project.

   HOW TO USE:
   1. Run  setupCompanyUpdatesTab()      → creates the tab.
   2. Select a company in Company_Master, run  addCompanyUpdate()
      → prompts for type / description / source and appends a dated entry.
   Programmatic API (for other units): logCompanyUpdate(id, type, desc, source).
   ============================================================================ */


// ── Tab name ─────────────────────────────────────────────────────────────────
const SHEET_UPDATES = { MAIN: "Company_Updates" };

// ── Column map (1-based) ──────────────────────────────────────────────────────
const COL_UPDATES = {
  COMPANY_ID  : 1,
  DATE        : 2,
  UPDATE_TYPE : 3,
  DESCRIPTION : 4,
  SOURCE      : 5
};

const UPDATE_TYPE_ENUM = ["News", "Legal", "Financial", "Leadership", "Correction", "Data"];

const INTEL_HEADER_BG_UP = "#1a1a2e";
const INTEL_HEADER_FG_UP = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupCompanyUpdatesTab() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  let up = ss.getSheetByName(SHEET_UPDATES.MAIN);
  const isNew = !up;
  if (isNew) up = ss.insertSheet(SHEET_UPDATES.MAIN);

  up.getRange(1, 1, 1, 5).setValues([["Company_ID", "Date", "Update_Type", "Description", "Source"]])
    .setBackground(INTEL_HEADER_BG_UP).setFontColor(INTEL_HEADER_FG_UP).setFontWeight("bold");
  up.setFrozenRows(1);
  [160, 110, 130, 460, 260].forEach((w, i) => up.setColumnWidth(i + 1, w));

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(UPDATE_TYPE_ENUM, true).setAllowInvalid(true).build();
  up.getRange(2, COL_UPDATES.UPDATE_TYPE, 999, 1).setDataValidation(rule);

  ui.alert("✅ Living Database Ready",
    (isNew ? "Created Company_Updates.\n\n" : "Company_Updates refreshed.\n\n") +
    "Add entries with 'Add Update — selected company', or programmatically via " +
    "logCompanyUpdate(id, type, description, source).",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// PROGRAMMATIC API — append one update (used by other units too)
// Also bumps Company_Master.Updated_At so the master row reflects the change.
// ══════════════════════════════════════════════════════════════════════════════
function logCompanyUpdate(companyId, updateType, description, source) {
  const ss = intelSS_();
  const up = ss.getSheetByName(SHEET_UPDATES.MAIN);
  if (!up) throw new Error("Company_Updates not found — run setupCompanyUpdatesTab() first.");

  up.appendRow([companyId, new Date(), updateType || "News", description || "", source || ""]);
  touchCompanyUpdatedAt_(companyId);
}

// Bump the Updated_At timestamp on the Company_Master row
function touchCompanyUpdatedAt_(companyId) {
  const ss = intelSS_();
  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  if (!master) return;
  const row = findCompanyRow_(master, companyId);   // from Intelligence_1.1
  if (row !== -1) master.getRange(row, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());
}


// ══════════════════════════════════════════════════════════════════════════════
// INTERACTIVE — add an update for the selected company (menu-friendly)
// ══════════════════════════════════════════════════════════════════════════════
function addCompanyUpdate() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  const companyId = getSelectedCompanyId_();   // from Intelligence_1.1
  if (!companyId) return;

  if (!ss.getSheetByName(SHEET_UPDATES.MAIN)) { ui.alert("Run setupCompanyUpdatesTab() first."); return; }

  // Type
  const typeResp = ui.prompt("Add Update — Type",
    "Enter update type (" + UPDATE_TYPE_ENUM.join(" / ") + "):", ui.ButtonSet.OK_CANCEL);
  if (typeResp.getSelectedButton() !== ui.Button.OK) return;
  let type = typeResp.getResponseText().trim();
  if (!type) type = "News";

  // Description
  const descResp = ui.prompt("Add Update — Description",
    "What happened? (one line)", ui.ButtonSet.OK_CANCEL);
  if (descResp.getSelectedButton() !== ui.Button.OK) return;
  const description = descResp.getResponseText().trim();
  if (!description) { ui.alert("No description entered — cancelled."); return; }

  // Source
  const srcResp = ui.prompt("Add Update — Source",
    "Source (URL or publication, optional):", ui.ButtonSet.OK_CANCEL);
  if (srcResp.getSelectedButton() !== ui.Button.OK) return;
  const source = srcResp.getResponseText().trim();

  logCompanyUpdate(companyId, type, description, source);
  ui.alert("✅ Update logged for " + companyId + ".");
}
