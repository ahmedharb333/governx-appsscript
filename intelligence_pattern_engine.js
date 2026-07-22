/* ============================================================================
   Intelligence_1.5_PatternEngine.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.5 — Pattern Engine

   Named failure patterns + a detector that scores how strongly each company
   exhibits each pattern, using its Company_Failure_Map + Company_DNA data.

   Tabs created:
     • Failure_Patterns   (the library — PT-###, seeded with 4 patterns)
     • Pattern_Company_Map(Pattern_ID × Company_ID × Weight 0-100)

   DETECTION IS ALGORITHMIC (no Claude call, no cost):
     weight = 0.6 · failure-signature match  +  0.4 · DNA-signature match
     - failure match = severity-weighted overlap with the pattern's signature FLs
     - DNA match     = fraction of the pattern's DNA conditions the company meets
   Companies scoring >= 20 are written to Pattern_Company_Map.

   SAFETY:
   - Additive. Own constants (SHEET_PATTERN, COL_PATTERN_*). No config.gs edits.
   - Reuses helpers from Intelligence_1.1 (normalizeName_, findCompanyRow_,
     getSelectedCompanyId_, logError) + reads Company_Failure_Map (1.3) and
     Company_DNA (1.4). Keep those unit files in the project.

   HOW TO USE:
   1. Run  setupPatternEngineTabs()  → creates + seeds 4 patterns and pre-maps
      their example companies (where found in Company_Master).
   2. Run  detectPatterns()          → scores the selected company vs all patterns.
      Or  detectAllPatterns()        → scores every company at once (cheap).
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_PATTERN = {
  LIBRARY : "Failure_Patterns",
  MAP     : "Pattern_Company_Map"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_PATTERN_LIB = {
  PATTERN_ID         : 1,
  PATTERN_NAME       : 2,
  DESCRIPTION        : 3,
  SIGNATURE_FAILURES : 4,   // comma list of FL-### that define the pattern
  SIGNATURE_DNA      : 5     // condition string, e.g. "Risk_Appetite>=8;Board_Independence<=3"
};

const COL_PATTERN_MAP = {
  PATTERN_ID : 1,
  COMPANY_ID : 2,
  WEIGHT     : 3
};

// ── Trait-name → Company_DNA column index (for parsing Signature_DNA) ─────────
const DNA_TRAIT_COL = {
  "decision_speed"         : COL_DNA.DECISION_SPEED,
  "risk_appetite"          : COL_DNA.RISK_APPETITE,
  "innovation"             : COL_DNA.INNOVATION,
  "centralization"         : COL_DNA.CENTRALIZATION,
  "board_independence"     : COL_DNA.BOARD_INDEPENDENCE,
  "compliance_maturity"    : COL_DNA.COMPLIANCE_MATURITY,
  "operational_complexity" : COL_DNA.OPERATIONAL_COMPLEXITY,
  "transparency"           : COL_DNA.TRANSPARENCY
};

// ── Seed patterns: [Name, Description, Signature_Failures, Signature_DNA, [examples]]
const PATTERN_SEED = [
  ["Founder Syndrome",
   "A dominant founder concentrates power, captures the board, and outruns oversight.",
   "FL-002, FL-005, FL-008, FL-034, FL-035",
   "Risk_Appetite>=8; Board_Independence<=3; Centralization>=8; Transparency<=4",
   ["WeWork", "FTX", "Theranos", "Abraaj"]],

  ["Technology Arrogance",
   "A dominant incumbent dismisses disruption and protects its legacy business too long.",
   "FL-025, FL-026, FL-027, FL-029, FL-030",
   "Innovation<=4; Decision_Speed<=4; Centralization>=6; Risk_Appetite<=5",
   ["Kodak", "Nokia", "BlackBerry", "Yahoo"]],

  ["Sales Culture",
   "Aggressive targets and incentives normalize misconduct until regulators arrive.",
   "FL-021, FL-024, FL-023, FL-017, FL-020",
   "Risk_Appetite>=7; Compliance_Maturity<=4; Transparency<=4",
   ["Wells Fargo", "Volkswagen", "Boeing"]],

  ["Unchecked Growth",
   "Explosive growth outpaces risk controls, hiding fragility until sudden collapse.",
   "FL-015, FL-013, FL-014, FL-040, FL-016",
   "Risk_Appetite>=9; Operational_Complexity>=7; Board_Independence<=4; Compliance_Maturity<=4",
   ["Enron", "Lehman", "Credit Suisse", "FTX"]]
];

const INTEL_HEADER_BG_PT = "#1a1a2e";
const INTEL_HEADER_FG_PT = "#ffffff";
const PATTERN_MIN_WEIGHT  = 20;   // minimum score to record a pattern match


// ══════════════════════════════════════════════════════════════════════════════
// SETUP — create + seed the 2 tabs, and pre-map example companies where found
// ══════════════════════════════════════════════════════════════════════════════
function setupPatternEngineTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── Failure_Patterns ──────────────────────────────────────────────────────
  let lib = ss.getSheetByName(SHEET_PATTERN.LIBRARY);
  const libIsNew = !lib;
  if (libIsNew) lib = ss.insertSheet(SHEET_PATTERN.LIBRARY);

  lib.getRange(1, 1, 1, 5).setValues([[
    "Pattern_ID", "Pattern_Name", "Description", "Signature_Failures", "Signature_DNA"
  ]]).setBackground(INTEL_HEADER_BG_PT).setFontColor(INTEL_HEADER_FG_PT).setFontWeight("bold");
  lib.setFrozenRows(1);
  [120, 190, 420, 260, 360].forEach((w, i) => lib.setColumnWidth(i + 1, w));

  let seeded = 0;
  if (lib.getLastRow() < 2) {
    const rows = PATTERN_SEED.map((p, i) => [
      "PT-" + String(i + 1).padStart(3, "0"), p[0], p[1], p[2], p[3]
    ]);
    lib.getRange(2, 1, rows.length, 5).setValues(rows);
    seeded = rows.length;
  }

  // ── Pattern_Company_Map ───────────────────────────────────────────────────
  let map = ss.getSheetByName(SHEET_PATTERN.MAP);
  const mapIsNew = !map;
  if (mapIsNew) map = ss.insertSheet(SHEET_PATTERN.MAP);

  map.getRange(1, 1, 1, 3).setValues([["Pattern_ID", "Company_ID", "Weight"]])
     .setBackground(INTEL_HEADER_BG_PT).setFontColor(INTEL_HEADER_FG_PT).setFontWeight("bold");
  map.setFrozenRows(1);
  [120, 160, 100].forEach((w, i) => map.setColumnWidth(i + 1, w));

  // ── Pre-map seed example companies (curated, weight 90) where they exist ──
  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  let preMapped = 0, notFound = [];
  if (master) {
    PATTERN_SEED.forEach((p, i) => {
      const patternId = "PT-" + String(i + 1).padStart(3, "0");
      p[4].forEach(exampleName => {
        const coId = findCompanyIdByName_(master, exampleName);
        if (!coId) { notFound.push(exampleName); return; }
        if (patternMapHas_(map, patternId, coId)) return; // avoid dup on rerun
        map.appendRow([patternId, coId, 90]);
        preMapped++;
      });
    });
  }

  ui.alert(
    "✅ Pattern Engine Ready",
    (seeded ? "Seeded " + seeded + " patterns.\n" : "Patterns already present.\n") +
    "Pre-mapped " + preMapped + " example companies." +
    (notFound.length ? "\n\nNot found in Company_Master (skipped): " + notFound.join(", ") : "") +
    "\n\nNext: run detectAllPatterns() to score every company, or detectPatterns() for one.",
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// DETECT — score one company against every pattern (algorithmic)
// ══════════════════════════════════════════════════════════════════════════════
function detectPatterns(companyId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const result = scoreCompanyAgainstPatterns_(ss, companyId);
  if (result === null) { ui.alert("Run setupPatternEngineTabs() first."); return; }

  writePatternMatches_(ss, companyId, result.matches);
  ui.alert(
    "✅ Patterns detected for " + companyId,
    result.matches.length
      ? result.matches.map(m => m.patternId + " (" + m.name + "): " + m.weight).join("\n")
      : "No pattern scored above the threshold (" + PATTERN_MIN_WEIGHT + ").\n" +
        "Make sure this company has Failure Map (1.3) and/or DNA (1.4) data.",
    ui.ButtonSet.OK
  );
}

// ── DETECT ALL — score every company in Company_Master (cheap, no API) ────────
function detectAllPatterns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  if (!master) { ui.alert("Company_Master not found."); return; }

  const ids = master.getRange(2, COL_COMPANY_MASTER.COMPANY_ID, Math.max(master.getLastRow() - 1, 0), 1)
                    .getValues().map(r => (r[0] || "").toString().trim()).filter(Boolean);

  let companies = 0, links = 0;
  ids.forEach(id => {
    const res = scoreCompanyAgainstPatterns_(ss, id);
    if (!res) return;
    writePatternMatches_(ss, id, res.matches);
    if (res.matches.length) { companies++; links += res.matches.length; }
  });

  ui.alert("✅ Pattern detection complete",
    "Scored " + ids.length + " companies.\n" +
    links + " pattern links written across " + companies + " companies.",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// CORE SCORING (shared)
// ══════════════════════════════════════════════════════════════════════════════
function scoreCompanyAgainstPatterns_(ss, companyId) {
  const lib = ss.getSheetByName(SHEET_PATTERN.LIBRARY);
  if (!lib || lib.getLastRow() < 2) return null;

  // Company's failures: { FL-### : severity }
  const failures = {};
  const fmap = ss.getSheetByName(SHEET_FAILURE.MAP);
  if (fmap) {
    fmap.getDataRange().getValues().slice(1).forEach(r => {
      if ((r[COL_FAILURE_MAP.COMPANY_ID - 1] || "").toString().trim() === companyId) {
        const fid = (r[COL_FAILURE_MAP.FAILURE_ID - 1] || "").toString().trim().toUpperCase();
        const sev = parseInt(r[COL_FAILURE_MAP.SEVERITY - 1], 10) || 5;
        if (fid) failures[fid] = sev;
      }
    });
  }

  // Company's DNA: { trait_lower : value }
  let dna = null;
  const dnaSheet = ss.getSheetByName(SHEET_DNA.MAIN);
  if (dnaSheet) {
    const row = dnaSheet.getDataRange().getValues().find(
      r => (r[COL_DNA.COMPANY_ID - 1] || "").toString().trim() === companyId);
    if (row) {
      dna = {};
      Object.keys(DNA_TRAIT_COL).forEach(t => {
        const v = parseInt(row[DNA_TRAIT_COL[t] - 1], 10);
        if (!isNaN(v)) dna[t] = v;
      });
    }
  }

  const patterns = lib.getRange(2, 1, lib.getLastRow() - 1, 5).getValues();
  const matches = [];

  patterns.forEach(p => {
    const patternId = (p[COL_PATTERN_LIB.PATTERN_ID - 1] || "").toString().trim();
    const name      = (p[COL_PATTERN_LIB.PATTERN_NAME - 1] || "").toString().trim();
    const sigFails  = (p[COL_PATTERN_LIB.SIGNATURE_FAILURES - 1] || "").toString()
                        .split(",").map(s => s.trim().toUpperCase()).filter(Boolean);
    const dnaConds  = parseDnaConditions_(p[COL_PATTERN_LIB.SIGNATURE_DNA - 1]);

    // Failure-signature match: severity-weighted overlap (0..1)
    let failureScore = 0;
    if (sigFails.length) {
      let sevSum = 0;
      sigFails.forEach(fid => { if (failures[fid]) sevSum += failures[fid]; });
      failureScore = sevSum / (10 * sigFails.length);
    }

    // DNA-signature match: fraction of conditions satisfied (0..1)
    let dnaScore = 0, dnaApplicable = false;
    if (dna && dnaConds.length) {
      dnaApplicable = true;
      let met = 0;
      dnaConds.forEach(c => { if (conditionMet_(dna[c.trait], c.op, c.value)) met++; });
      dnaScore = met / dnaConds.length;
    }

    // Combine — if no DNA yet, rely on failures only (and vice versa)
    let weight;
    if (dnaApplicable && sigFails.length) weight = 0.6 * failureScore + 0.4 * dnaScore;
    else if (dnaApplicable)               weight = dnaScore;
    else                                  weight = failureScore;

    const w = Math.round(weight * 100);
    if (w >= PATTERN_MIN_WEIGHT) matches.push({ patternId, name, weight: w });
  });

  matches.sort((a, b) => b.weight - a.weight);
  return { matches };
}

// Rewrite this company's rows in Pattern_Company_Map.
// Only overwrites when we actually computed matches — so running detection on a
// company that has no failure/DNA data yet leaves its curated seed rows intact.
function writePatternMatches_(ss, companyId, matches) {
  const map = ss.getSheetByName(SHEET_PATTERN.MAP);
  if (!map || !matches.length) return;

  const data = map.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][COL_PATTERN_MAP.COMPANY_ID - 1] || "").toString().trim() === companyId) {
      map.deleteRow(i + 1);
    }
  }
  const rows = matches.map(m => [m.patternId, companyId, m.weight]);
  map.getRange(map.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

// Parse "Risk_Appetite>=8; Board_Independence<=3" → [{trait, op, value}]
function parseDnaConditions_(str) {
  const out = [];
  (str || "").toString().split(/[;,]/).forEach(part => {
    const m = part.trim().match(/([A-Za-z_]+)\s*(>=|<=|>|<|=)\s*(\d+)/);
    if (m) out.push({ trait: m[1].toLowerCase(), op: m[2], value: parseInt(m[3], 10) });
  });
  return out;
}

function conditionMet_(val, op, target) {
  if (val === undefined || val === null || isNaN(val)) return false;
  switch (op) {
    case ">=": return val >= target;
    case "<=": return val <= target;
    case ">":  return val >  target;
    case "<":  return val <  target;
    case "=":  return val === target;
  }
  return false;
}

// Find a Company_ID by (normalized) name
function findCompanyIdByName_(master, name) {
  const target = normalizeName_(name);   // from Intelligence_1.1
  const data = master.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeName_(data[i][COL_COMPANY_MASTER.COMPANY_NAME - 1] || "") === target) {
      return (data[i][COL_COMPANY_MASTER.COMPANY_ID - 1] || "").toString().trim();
    }
  }
  return "";
}

function patternMapHas_(map, patternId, companyId) {
  const data = map.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_PATTERN_MAP.PATTERN_ID - 1] || "").toString().trim() === patternId &&
        (data[i][COL_PATTERN_MAP.COMPANY_ID - 1] || "").toString().trim() === companyId) return true;
  }
  return false;
}
