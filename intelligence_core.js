/* ============================================================================
   Intelligence_Core.gs — GovernX Intelligence Platform
   Shared foundation for the intelligence layer.

   PURPOSE:
   1. intelSS_()  — single access point for the intelligence spreadsheet.
      Today it returns the active (production) spreadsheet, so everything lives
      in one file. If you later split intelligence into its own spreadsheet,
      set INTEL_SPREADSHEET_ID once here and every unit follows — no rewrite.
   2. organizeIntelligenceTabs() — color-codes all intelligence tabs navy and
      groups them at the end of the tab bar, visually separating them from the
      production pipeline tabs.

   SAFETY: additive, own constants, no config.gs edits.
   ============================================================================ */


// ── Where the intelligence tabs live ─────────────────────────────────────────
// "" (blank)  → use the active/production spreadsheet (current: one-sheet mode).
// "<id>"      → use a separate spreadsheet by ID (future split — one change here).
const INTEL_SPREADSHEET_ID = "";

function intelSS_() {
  return INTEL_SPREADSHEET_ID
    ? SpreadsheetApp.openById(INTEL_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}


// ── Visual grouping ──────────────────────────────────────────────────────────
const INTEL_TAB_COLOR = "#1F3A5F";  // navy — marks a tab as "intelligence"

// All intelligence tabs, in the order they should appear (existing + planned).
// Only tabs that actually exist are touched.
const INTEL_TABS = [
  "Company_Master", "Company_Profile", "Company_Timeline", "Company_Relationships",
  "Failure_Taxonomy", "Company_Failure_Map", "Company_DNA",
  "Failure_Patterns", "Pattern_Company_Map",
  "Knowledge_Graph", "Graph_Edges", "Company_Updates",
  "GovernX_Framework", "Framework_Assessment", "Governance_Scorecard",
  "MrX_Library", "Character_Blueprint", "MrX_Renders",
  "Affiliate_Assets", "Digital_Products", "Audience_Memory",
  "Intelligence_Dashboard",
  "Research_Sources", "Research_Claims", "Research_Conflicts", "Research_Data_Moments"
];

// Color-code + move every existing intelligence tab to the end, in order.
function organizeIntelligenceTabs() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();
  let touched = 0;

  INTEL_TABS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setTabColor(INTEL_TAB_COLOR);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(ss.getNumSheets());  // push to the far right, in list order
    touched++;
  });

  ui.alert(
    "✅ Intelligence Tabs Organized",
    "Colored + moved " + touched + " intelligence tab(s) to the end of the tab bar.\n\n" +
    "Note: Google Sheets has no browser-style tab groups — color + contiguous order " +
    "is the closest it allows. Use Hide/Show below to collapse them like a group.",
    ui.ButtonSet.OK
  );
}

// ── Collapse / expand: the closest thing Sheets has to a browser tab group ────
// Hide every intelligence tab (declutter while working on production).
function hideIntelligenceTabs() {
  const ss = intelSS_();
  let hidden = 0;
  INTEL_TABS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) { sh.hideSheet(); hidden++; }
  });
  SpreadsheetApp.getUi().alert("🙈 Intelligence tabs hidden: " + hidden +
    "\n\nRun 'Show Intelligence Tabs' to bring them back.");
}

// Reveal every intelligence tab again.
function showIntelligenceTabs() {
  const ss = intelSS_();
  let shown = 0;
  INTEL_TABS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.isSheetHidden()) { sh.showSheet(); shown++; }
  });
  SpreadsheetApp.getUi().alert("👁️ Intelligence tabs shown: " + shown);
}
