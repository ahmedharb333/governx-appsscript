/* ============================================================================
   Intelligence_3.1_MrXCharacter.gs — GovernX Intelligence Platform
   PHASE 3 · UNIT 3.1 — Mr. X Character Assets

   Mr. X is the GovernX on-screen host/analyst.
   CANON (locked): professional male analyst · round black glasses · full beard ·
   dark hair · navy suit · red tie · slight smile.

   Tabs created:
     • MrX_Library       (the reusable expression set — 6 seeded expressions)
     • Character_Blueprint(per-scene direction for a given video)

   This unit stores the ASSETS + the per-scene plan. The lip-sync animation
   itself (Stage 7C, D-ID / HeyGen) comes in Unit 3.2. You can build/test the
   blueprint now with a placeholder PNG and wire the vendor later.

   SAFETY: additive, own constants, intelSS_(), no config.gs edits.
   Tabs are pre-registered in INTEL_TABS → covered by organize/hide/show.

   HOW TO USE:
   1. Run  setupMrXTabs()  → creates both tabs, seeds the 6 expressions.
   2. Generate the 6 expression PNGs (transparent background) once, upload to
      Drive, and paste each URL into the MrX_Library PNG column.
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_MRX = {
  LIBRARY   : "MrX_Library",
  BLUEPRINT : "Character_Blueprint"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_MRX_LIB = {
  EXPRESSION : 1,
  PNG        : 2,   // Drive URL — transparent-bg portrait
  ANIMATION  : 3,   // idle/nod/point/lean-in … or Drive URL to a clip
  USAGE      : 4
};

const COL_MRX_BLUEPRINT = {
  SCENE_ID   : 1,
  EXPRESSION : 2,
  GESTURE    : 3,
  EMOTION    : 4,
  POSITION   : 5,
  ANIMATION  : 6
};

// ── Controlled vocabularies ──────────────────────────────────────────────────
const MRX_EXPRESSIONS = ["Neutral", "Serious", "Thinking", "Warning", "Surprised", "Happy"];
const MRX_POSITIONS   = ["corner-BR", "corner-BL", "center", "lower-third"];

// ── Seed: [Expression, Usage] (PNG + Animation filled after art is produced) ─
const MRX_SEED = [
  ["Neutral",   "Default narration — opening titles, transitions, neutral exposition."],
  ["Serious",   "Collapse moments, root-cause reveals, grave statistics, the GRC verdict."],
  ["Thinking",  "Timeline/analysis segments; posing the 'how did this happen?' question."],
  ["Warning",   "Counter animations, risk spikes, the beat right before the fall."],
  ["Surprised", "Shocking data reveals, plot-twist facts, unexpected figures."],
  ["Happy",     "Positive turnarounds, the lesson resolution, the CTA / subscribe outro."]
];

const INTEL_HEADER_BG_MX = "#1a1a2e";
const INTEL_HEADER_FG_MX = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupMrXTabs() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  // ── MrX_Library ───────────────────────────────────────────────────────────
  let lib = ss.getSheetByName(SHEET_MRX.LIBRARY);
  const libNew = !lib;
  if (libNew) lib = ss.insertSheet(SHEET_MRX.LIBRARY);
  lib.getRange(1, 1, 1, 4).setValues([["Expression", "PNG", "Animation", "Usage"]])
     .setBackground(INTEL_HEADER_BG_MX).setFontColor(INTEL_HEADER_FG_MX).setFontWeight("bold");
  lib.setFrozenRows(1);
  [130, 300, 200, 480].forEach((w, i) => lib.setColumnWidth(i + 1, w));

  let seeded = 0;
  if (lib.getLastRow() < 2) {
    const rows = MRX_SEED.map(s => [s[0], "", "", s[1]]);
    lib.getRange(2, 1, rows.length, 4).setValues(rows);
    seeded = rows.length;
  }
  const exprRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MRX_EXPRESSIONS, true).setAllowInvalid(true).build();
  lib.getRange(2, COL_MRX_LIB.EXPRESSION, Math.max(lib.getLastRow() - 1, 999), 1).setDataValidation(exprRule);

  // ── Character_Blueprint ───────────────────────────────────────────────────
  let bp = ss.getSheetByName(SHEET_MRX.BLUEPRINT);
  const bpNew = !bp;
  if (bpNew) bp = ss.insertSheet(SHEET_MRX.BLUEPRINT);
  bp.getRange(1, 1, 1, 6).setValues([["Scene_ID", "Expression", "Gesture", "Emotion", "Position", "Animation"]])
    .setBackground(INTEL_HEADER_BG_MX).setFontColor(INTEL_HEADER_FG_MX).setFontWeight("bold");
  bp.setFrozenRows(1);
  [200, 130, 220, 160, 130, 200].forEach((w, i) => bp.setColumnWidth(i + 1, w));
  bp.getRange(2, COL_MRX_BLUEPRINT.EXPRESSION, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MRX_EXPRESSIONS, true).setAllowInvalid(true).build());
  bp.getRange(2, COL_MRX_BLUEPRINT.POSITION, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MRX_POSITIONS, true).setAllowInvalid(true).build());

  ui.alert("✅ Mr. X Character Ready",
    (seeded ? "Seeded " + seeded + " expressions in MrX_Library.\n" : "MrX_Library already populated.\n") +
    (bpNew ? "Created Character_Blueprint.\n\n" : "Character_Blueprint refreshed.\n\n") +
    "Next: generate the 6 expression PNGs (transparent bg), upload to Drive, and " +
    "paste each URL into the MrX_Library PNG column.\n\n" +
    "Stage 7C (lip-sync via D-ID/HeyGen) arrives in Unit 3.2.",
    ui.ButtonSet.OK);
}

// Quick check: are all 6 expression PNGs filled in?
function checkMrXArtReady() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();
  const lib = ss.getSheetByName(SHEET_MRX.LIBRARY);
  if (!lib) { ui.alert("Run setupMrXTabs() first."); return; }

  const data = lib.getDataRange().getValues();
  const missing = [];
  for (let i = 1; i < data.length; i++) {
    const expr = (data[i][COL_MRX_LIB.EXPRESSION - 1] || "").toString().trim();
    const png  = (data[i][COL_MRX_LIB.PNG - 1] || "").toString().trim();
    if (expr && !png) missing.push(expr);
  }
  ui.alert(missing.length ? "⚠️ Missing PNGs" : "✅ All expression art present",
    missing.length ? "No PNG yet for: " + missing.join(", ") : "All 6 expressions have a PNG URL.",
    ui.ButtonSet.OK);
}
