/* ============================================================================
   Intelligence_5.1_AudienceMemory.gs — GovernX Intelligence Platform
   PHASE 5 · UNIT 5.1 — Audience Memory (the learning engine)

   Extends Channel Memory: correlates YouTube performance (CTR / retention /
   views) against each video's attributes (Mr. X expression mix, scene types,
   CTA) and learns which choices perform best — then feeds that back into future
   generation via readAudienceMemory().

   Tab created:
     • Audience_Memory (Metric, Value, Confidence, Sample_Size, Updated_At)

   SAFETY: additive, own constants, intelSS_(), no config.gs edits.
   Reads production Publishing Tracker (perf) + Visual Library + Script Bank, and
   the intelligence Character_Blueprint. Reuses dominantExpression_ (Unit 3.2).

   HOW TO USE:
   1. Run  setupAudienceMemoryTab()  → creates + seeds the metric rows.
   2. Run  updateAudienceMemory()     → learns from videos that have perf data.
      (Ideally call this from the weekly YouTube Analytics job.)
   3. readAudienceMemory() returns a string other stages can inject into prompts.
   ============================================================================ */


// ── Tab name ─────────────────────────────────────────────────────────────────
const SHEET_AUDIENCE = { MAIN: "Audience_Memory" };

// ── Column map (1-based) ──────────────────────────────────────────────────────
const COL_AUDIENCE = {
  METRIC      : 1,
  VALUE       : 2,
  CONFIDENCE  : 3,
  SAMPLE_SIZE : 4,
  UPDATED_AT  : 5
};

// Metrics we track. (Expression / Scene_Type / CTA are auto-learned; the others
// are seeded for you to fill manually or extend later.)
const AUDIENCE_METRICS = [
  "Best_Expression", "Best_Scene_Type", "Best_CTA",
  "Best_Hook", "Best_Color", "Best_Thumbnail"
];

const INTEL_HEADER_BG_AM = "#1a1a2e";
const INTEL_HEADER_FG_AM = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupAudienceMemoryTab() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  let am = ss.getSheetByName(SHEET_AUDIENCE.MAIN);
  const isNew = !am;
  if (isNew) am = ss.insertSheet(SHEET_AUDIENCE.MAIN);

  am.getRange(1, 1, 1, 5).setValues([["Metric", "Value", "Confidence", "Sample_Size", "Updated_At"]])
    .setBackground(INTEL_HEADER_BG_AM).setFontColor(INTEL_HEADER_FG_AM).setFontWeight("bold");
  am.setFrozenRows(1);
  [200, 260, 120, 120, 160].forEach((w, i) => am.setColumnWidth(i + 1, w));

  // Seed metric rows if empty
  if (am.getLastRow() < 2) {
    const rows = AUDIENCE_METRICS.map(m => [m, "", "", "", ""]);
    am.getRange(2, 1, rows.length, 5).setValues(rows);
  }

  ui.alert("✅ Audience Memory Ready",
    (isNew ? "Created Audience_Memory with " + AUDIENCE_METRICS.length + " metrics.\n\n"
           : "Audience_Memory refreshed.\n\n") +
    "Run updateAudienceMemory() once videos have CTR/Retention data to learn " +
    "Best_Expression, Best_Scene_Type, and Best_CTA automatically.",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// LEARN — correlate performance with attributes and update the winners
// ══════════════════════════════════════════════════════════════════════════════
function updateAudienceMemory() {
  const prodSS = SpreadsheetApp.getActiveSpreadsheet();
  const ss     = intelSS_();
  const ui     = SpreadsheetApp.getUi();

  const am = ss.getSheetByName(SHEET_AUDIENCE.MAIN);
  if (!am) { ui.alert("Run setupAudienceMemoryTab() first."); return; }

  // ── Gather videos that have performance data ──────────────────────────────
  const videos = getPerformanceVideos_(prodSS);
  if (videos.length === 0) {
    ui.alert("No performance data yet",
      "No videos in Publishing Tracker have CTR/Retention filled.\n" +
      "Run the YouTube Analytics pull first, then rerun this.", ui.ButtonSet.OK);
    return;
  }

  // ── Learn each auto-metric ────────────────────────────────────────────────
  const results = {};
  results["Best_Expression"] = learnBestAttribute_(videos, v => dominantExpression_(ss, v.videoId));
  results["Best_Scene_Type"] = learnBestAttribute_(videos, v => dominantSceneType_(prodSS, v.videoId));
  results["Best_CTA"]        = learnBestAttribute_(videos, v => videoCTA_(prodSS, v.videoId));

  // ── Write winners back ────────────────────────────────────────────────────
  let updated = 0;
  Object.keys(results).forEach(metric => {
    const r = results[metric];
    if (!r) return;
    const row = findAudienceRow_(am, metric);
    if (row === -1) return;
    am.getRange(row, COL_AUDIENCE.VALUE      ).setValue(r.value);
    am.getRange(row, COL_AUDIENCE.CONFIDENCE ).setValue(r.confidence);
    am.getRange(row, COL_AUDIENCE.SAMPLE_SIZE).setValue(r.sampleSize);
    am.getRange(row, COL_AUDIENCE.UPDATED_AT ).setValue(new Date());
    updated++;
  });

  ui.alert("✅ Audience Memory updated",
    "Learned from " + videos.length + " video(s) with performance data.\n\n" +
    "Best_Expression: " + fmtResult_(results["Best_Expression"]) + "\n" +
    "Best_Scene_Type: " + fmtResult_(results["Best_Scene_Type"]) + "\n" +
    "Best_CTA: "        + fmtResult_(results["Best_CTA"]) + "\n\n" +
    (videos.length < 3 ? "⚠️ Confidence is low until you have several published videos." : ""),
    ui.ButtonSet.OK);
}

function fmtResult_(r) {
  return r ? (r.value + " (conf " + r.confidence + ", n=" + r.sampleSize + ")") : "— (no data)";
}


// ══════════════════════════════════════════════════════════════════════════════
// READ — a compact string other stages can inject into their prompts
// (like Channel Memory). Only includes confident-enough signals.
// ══════════════════════════════════════════════════════════════════════════════
function readAudienceMemory() {
  const ss = intelSS_();
  const am = ss.getSheetByName(SHEET_AUDIENCE.MAIN);
  if (!am) return "";
  const data = am.getDataRange().getValues();
  const lines = [];
  for (let i = 1; i < data.length; i++) {
    const metric = data[i][COL_AUDIENCE.METRIC - 1];
    const value  = data[i][COL_AUDIENCE.VALUE - 1];
    const conf   = Number(data[i][COL_AUDIENCE.CONFIDENCE - 1]) || 0;
    if (value && conf >= 40) lines.push(metric + ": " + value + " (confidence " + conf + ")");
  }
  if (!lines.length) return "";
  return "AUDIENCE MEMORY — what has performed best (bias new content toward these):\n" + lines.join("\n");
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
// Videos with a computable performance score
function getPerformanceVideos_(prodSS) {
  const pub = prodSS.getSheetByName(SHEET.PUBLISHING);
  if (!pub) return [];
  const data = pub.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const id  = (data[i][COL_PUBLISHING.ID - 1] || "").toString().trim();
    if (!id) continue;
    const ctr = Number(data[i][COL_PUBLISHING.CTR - 1]);
    const ret = Number(data[i][COL_PUBLISHING.RETENTION - 1]);
    if (isNaN(ctr) && isNaN(ret)) continue;   // no perf data
    // Composite score (CTR weighted, retention added). Relative ranking only.
    const perf = (isNaN(ctr) ? 0 : ctr * 4) + (isNaN(ret) ? 0 : ret);
    out.push({ videoId: id, perf: perf });
  }
  return out;
}

// Group videos by an attribute value, average perf, return the best
function learnBestAttribute_(videos, attrFn) {
  const groups = {};
  videos.forEach(v => {
    let a;
    try { a = attrFn(v); } catch (e) { a = null; }
    if (a === null || a === undefined || a === "") return;
    a = a.toString();
    if (!groups[a]) groups[a] = { sum: 0, n: 0 };
    groups[a].sum += v.perf; groups[a].n += 1;
  });
  const keys = Object.keys(groups);
  if (!keys.length) return null;

  let best = null, bestAvg = -Infinity;
  keys.forEach(k => {
    const avg = groups[k].sum / groups[k].n;
    if (avg > bestAvg) { bestAvg = avg; best = k; }
  });
  const n = groups[best].n;
  const confidence = Math.min(90, n * 25 + (keys.length > 1 ? 10 : 0));
  return { value: best, sampleSize: n, confidence: confidence };
}

// dominantExpression_ lives in Unit 3.2 (reads Character_Blueprint) — reused.

// Most common scene type for a video (from Visual Library)
function dominantSceneType_(prodSS, videoId) {
  const v = prodSS.getSheetByName(SHEET.VISUAL);
  if (!v) return null;
  const data = v.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_VISUAL.ID - 1] || "").toString().trim() !== videoId) continue;
    const t = (data[i][COL_VISUAL.SCENE_TYPE - 1] || "").toString().trim();
    if (t) counts[t] = (counts[t] || 0) + 1;
  }
  let best = null, max = 0;
  Object.keys(counts).forEach(k => { if (counts[k] > max) { max = counts[k]; best = k; } });
  return best;
}

// The video's call-to-action (from Script Bank)
function videoCTA_(prodSS, videoId) {
  const s = prodSS.getSheetByName(SHEET.SCRIPT);
  if (!s) return null;
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_SCRIPT.ID - 1] || "").toString().trim() === videoId) {
      return (data[i][COL_SCRIPT.CALL_TO_ACTION - 1] || "").toString().trim() || null;
    }
  }
  return null;
}

function findAudienceRow_(am, metric) {
  const data = am.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_AUDIENCE.METRIC - 1] || "").toString().trim() === metric) return i + 1;
  }
  return -1;
}
 