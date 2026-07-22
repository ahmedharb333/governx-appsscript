/* ============================================================================
   Video_Economics.gs — GovernX Content OS
   Per-video COST vs REVENUE, ROI, and break-even.

   COST is auto-estimated from what the pipeline ACTUALLY produced for each video:
     • ElevenLabs — exact: billed per character, and the narration character count
       is the script itself (Stage 7 full VO + Stage 7B per-scene = 2 passes).
     • Claude — estimated from the size of everything the model generated for the
       video (research + script + scene data + metadata) plus a per-stage input
       allowance. Approximate; the exact figure would need token logging inside
       callClaude (see NOTE at bottom).
     • Thumbnail — flat per image (0 when rendered in Remotion).

   REVENUE is pulled from the YouTube Analytics API (estimatedRevenue), which needs
   the monetary scope and a monetized channel — see pullEconomicsRevenue().

   Rates live in EDITABLE CELLS at the top of the tab, so you tune them without
   touching code. "🔄 Refresh Video Economics" recomputes every row.
   ============================================================================ */

const SHEET_ECON = "Video Economics";

// Default rates (written into the config block on first build; the tab's cells win)
const ECON_DEFAULTS = {
  claudeInPerM   : 3.00,   // $ per 1M input tokens  (Claude Sonnet-class blended)
  claudeOutPerM  : 15.00,  // $ per 1M output tokens
  charsPerToken  : 4,      // rough chars→tokens
  claudeInputAllowTokens: 9000,  // est. input tokens PER pipeline stage (system+prompt, cached)
  claudeStages   : 6,      // stages that call Claude (1,2,3,3B,4,4B,10 ≈ 6 billable after cache)
  elevenPer1k    : 0.30,   // $ per 1,000 characters
  elevenPasses   : 2,      // Stage 7 (full) + Stage 7B (per-scene) = 2× the script
  thumbnailFlat  : 0.00    // $ per thumbnail (0 = Remotion; set ~0.08 for Ideogram)
};

// Config-block cell addresses (label col A, value col B)
const ECON_CFG_ROWS = {
  claudeInPerM: 2, claudeOutPerM: 3, charsPerToken: 4,
  claudeInputAllowTokens: 5, claudeStages: 6,
  elevenPer1k: 7, elevenPasses: 8, thumbnailFlat: 9
};
const ECON_TABLE_TOP = 12;   // header row of the per-video table


function buildVideoEconomicsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ECON);
  if (!sh) sh = ss.insertSheet(SHEET_ECON);
  sh.clear();

  // ── Config block ──────────────────────────────────────────────────────────
  sh.getRange(1, 1).setValue("⚙️  RATES — edit these, then Refresh").setFontWeight("bold").setFontSize(12);
  const cfg = [
    ["Claude — $ / 1M input tokens",  ECON_DEFAULTS.claudeInPerM],
    ["Claude — $ / 1M output tokens", ECON_DEFAULTS.claudeOutPerM],
    ["Chars per token",               ECON_DEFAULTS.charsPerToken],
    ["Claude — input tokens / stage", ECON_DEFAULTS.claudeInputAllowTokens],
    ["Claude — billable stages",      ECON_DEFAULTS.claudeStages],
    ["ElevenLabs — $ / 1,000 chars",  ECON_DEFAULTS.elevenPer1k],
    ["ElevenLabs — passes (7 + 7B)",  ECON_DEFAULTS.elevenPasses],
    ["Thumbnail — $ / image",         ECON_DEFAULTS.thumbnailFlat]
  ];
  cfg.forEach(function (row, i) {
    sh.getRange(i + 2, 1).setValue(row[0]).setFontColor("#64748B");
    sh.getRange(i + 2, 2).setValue(row[1]).setFontWeight("bold").setBackground("#FEF3C7");
  });

  // ── Table header ──────────────────────────────────────────────────────────
  const headers = ["ID", "Company", "Title", "Published",
    "Claude $", "ElevenLabs $", "Thumbnail $", "TOTAL COST $",
    "Views", "Revenue $", "RPM $", "NET $", "ROI %", "Status"];
  headers.forEach(function (h, i) {
    sh.getRange(ECON_TABLE_TOP, i + 1).setValue(h).setFontWeight("bold")
      .setBackground("#1E293B").setFontColor("#FFFFFF").setFontFamily("Montserrat")
      .setHorizontalAlignment(i >= 4 ? "center" : "left");
  });
  [150, 200, 300, 100, 90, 100, 100, 110, 90, 100, 80, 100, 80, 120]
    .forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(ECON_TABLE_TOP);

  refreshVideoEconomics();
  SpreadsheetApp.getUi().alert("✅ Video Economics tab ready",
    "Rates are in the yellow cells — tune them and Refresh.\n\n" +
    "Cost is auto-estimated from each video's real script/research/scene sizes.\n" +
    "Run '📈 Pull Video Revenue' to fill Revenue from YouTube Analytics.",
    SpreadsheetApp.getUi().ButtonSet.OK);
}


function refreshVideoEconomics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ECON);
  if (!sh) { buildVideoEconomicsTab(); return; }

  const R = econReadRates_(sh);

  // Every video that has reached at least Stage 1 (Master Content).
  const master = ss.getSheetByName(SHEET.MASTER);
  const ids = [];
  if (master && master.getLastRow() > 1) {
    master.getRange(2, 1, master.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { const id = String(r[0]).trim(); if (id) ids.push(id); });
  }

  // Preserve any Revenue/Views already pulled (don't clobber on a cost refresh).
  const existing = econExistingRevenue_(sh);

  // Clear old body.
  const lastRow = sh.getLastRow();
  if (lastRow > ECON_TABLE_TOP) sh.getRange(ECON_TABLE_TOP + 1, 1, lastRow - ECON_TABLE_TOP, 14).clearContent();

  const rows = ids.map(function (id) {
    const cost = econEstimateCost_(ss, id, R);
    const rev  = existing[id] || { views: "", revenue: "" };
    return [id, cost.company, cost.title, cost.published,
      round2(cost.claude), round2(cost.eleven), round2(cost.thumb), round2(cost.total),
      rev.views, rev.revenue, "", "", "", ""];
  });

  if (!rows.length) return;
  sh.getRange(ECON_TABLE_TOP + 1, 1, rows.length, 14).setValues(rows);

  // Formula columns: RPM, NET, ROI, Status (reference the row's own cells).
  for (var i = 0; i < rows.length; i++) {
    const r = ECON_TABLE_TOP + 1 + i;
    sh.getRange(r, 11).setFormula('=IFERROR(IF(I' + r + '>0, J' + r + '/I' + r + '*1000, ""), "")');       // RPM
    sh.getRange(r, 12).setFormula('=IFERROR(IF(J' + r + '="","",J' + r + '-H' + r + '), "")');               // NET
    sh.getRange(r, 13).setFormula('=IFERROR(IF(AND(J' + r + '<>"",H' + r + '>0),(J' + r + '-H' + r + ')/H' + r + ',""), "")'); // ROI
    sh.getRange(r, 14).setFormula('=IF(J' + r + '="","⏳ no revenue yet",IF(J' + r + '>=H' + r + ',"✅ profit","🔻 loss"))');   // Status
  }
  sh.getRange(ECON_TABLE_TOP + 1, 5, rows.length, 8).setNumberFormat("$#,##0.00");
  sh.getRange(ECON_TABLE_TOP + 1, 13, rows.length, 1).setNumberFormat("0%");
  SpreadsheetApp.flush();
}


/* ── Cost estimate for one video, from what the pipeline produced ────────────── */
function econEstimateCost_(ss, id, R) {
  const out = { company: "", title: "", published: "", claude: 0, eleven: 0, thumb: 0, total: 0 };

  // Company + published from Master / Publishing.
  const master = ss.getSheetByName(SHEET.MASTER);
  if (master) {
    const md = master.getDataRange().getValues();
    for (var i = 1; i < md.length; i++) if (String(md[i][0]).trim() === id) { out.title = String(md[i][1] || ""); break; }
  }
  const idea = ss.getSheetByName(SHEET.IDEA);
  if (idea) {
    const d = idea.getDataRange().getValues();
    for (var j = 1; j < d.length; j++) if (String(d[j][0]).trim() === id) { out.company = String(d[j][COL_IDEA.COMPANY - 1] || ""); break; }
  }
  const pub = ss.getSheetByName(SHEET.PUBLISHING);
  if (pub) {
    const p = pub.getDataRange().getValues();
    for (var k = 1; k < p.length; k++) if (String(p[k][0]).trim() === id) {
      const dt = p[k][COL_PUBLISHING.PUBLISH_DATE - 1];
      out.published = dt ? (dt instanceof Date ? Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(dt)) : "";
      break;
    }
  }

  // ── Character counts of what the model generated + the narration script ─────
  const scriptChars   = econSumChars_(ss, SHEET.SCRIPT,   id, [COL_SCRIPT.VOICEOVER_SCRIPT, COL_SCRIPT.HOOK, COL_SCRIPT.NARRATIVE_FLOW, COL_SCRIPT.ANALYSIS_FRAMEWORK, COL_SCRIPT.GRC_BPR_CLOSING]);
  const voChars       = econSumChars_(ss, SHEET.SCRIPT,   id, [COL_SCRIPT.VOICEOVER_SCRIPT]);
  const researchChars = econSumChars_(ss, SHEET.RESEARCH, id, "ALL");
  const sceneChars    = econSumChars_(ss, SHEET.VISUAL,   id, [COL_VISUAL.DESCRIPTION, COL_VISUAL_EXTENDED.REMOTION_DATA, COL_VISUAL_EXTENDED.VOICEOVER_SYNC]);
  const metaChars     = econSumChars_(ss, SHEET.YOUTUBE,  id, "ALL");

  // Claude cost: output tokens from generated chars; input from a per-stage allowance.
  const outTokens = (scriptChars + researchChars + sceneChars + metaChars) / R.charsPerToken;
  const inTokens  = R.claudeInputAllowTokens * R.claudeStages;
  out.claude = (inTokens / 1e6) * R.claudeInPerM + (outTokens / 1e6) * R.claudeOutPerM;

  // ElevenLabs: script narration chars × passes (Stage 7 full + Stage 7B per-scene).
  out.eleven = (voChars * R.elevenPasses / 1000) * R.elevenPer1k;

  out.thumb = R.thumbnailFlat;
  out.total = out.claude + out.eleven + out.thumb;
  return out;
}

function econSumChars_(ss, sheetName, id, cols) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return 0;
  const data = sh.getDataRange().getValues();
  let total = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== id) continue;
    if (cols === "ALL") { data[i].forEach(function (c) { total += String(c || "").length; }); }
    else { cols.forEach(function (c) { if (c) total += String(data[i][c - 1] || "").length; }); }
  }
  return total;
}

function econReadRates_(sh) {
  const g = function (row, def) { const v = sh.getRange(row, 2).getValue(); return (v === "" || isNaN(v)) ? def : Number(v); };
  return {
    claudeInPerM:  g(ECON_CFG_ROWS.claudeInPerM,  ECON_DEFAULTS.claudeInPerM),
    claudeOutPerM: g(ECON_CFG_ROWS.claudeOutPerM, ECON_DEFAULTS.claudeOutPerM),
    charsPerToken: g(ECON_CFG_ROWS.charsPerToken, ECON_DEFAULTS.charsPerToken) || 4,
    claudeInputAllowTokens: g(ECON_CFG_ROWS.claudeInputAllowTokens, ECON_DEFAULTS.claudeInputAllowTokens),
    claudeStages:  g(ECON_CFG_ROWS.claudeStages,  ECON_DEFAULTS.claudeStages),
    elevenPer1k:   g(ECON_CFG_ROWS.elevenPer1k,   ECON_DEFAULTS.elevenPer1k),
    elevenPasses:  g(ECON_CFG_ROWS.elevenPasses,  ECON_DEFAULTS.elevenPasses),
    thumbnailFlat: g(ECON_CFG_ROWS.thumbnailFlat, ECON_DEFAULTS.thumbnailFlat)
  };
}

function econExistingRevenue_(sh) {
  const map = {};
  const last = sh.getLastRow();
  if (last <= ECON_TABLE_TOP) return map;
  const data = sh.getRange(ECON_TABLE_TOP + 1, 1, last - ECON_TABLE_TOP, 10).getValues();
  data.forEach(function (r) {
    const id = String(r[0]).trim();
    if (id) map[id] = { views: r[8], revenue: r[9] };
  });
  return map;
}

function round2(n) { return Math.round((Number(n) || 0) * 100) / 100; }


/* ── Revenue via the YouTube Analytics API (estimatedRevenue) ─────────────────
   Needs: Advanced Service "YouTubeAnalytics" enabled, a MONETIZED channel, and
   the monetary scope. Reuses the videoId in the Publishing Tracker.
   ---------------------------------------------------------------------------- */
function pullEconomicsRevenue() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ECON);
  if (!sh) { buildVideoEconomicsTab(); return; }

  if (typeof YouTubeAnalytics === "undefined") {
    ui.alert("YouTube Analytics service not enabled",
      "In the Apps Script editor: Services (＋) → add 'YouTube Analytics API' → save as 'YouTubeAnalytics'.\n\n" +
      "Your channel must be monetized for estimatedRevenue to return data.", ui.ButtonSet.OK);
    return;
  }

  const pub = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pub) { ui.alert("No Publishing Tracker."); return; }
  const pd = pub.getDataRange().getValues();

  // id → videoId + publishDate
  const vids = {};
  for (var i = 1; i < pd.length; i++) {
    const id  = String(pd[i][COL_PUBLISHING.ID - 1]).trim();
    const vid = String(pd[i][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1] || "").trim();
    if (id && vid.length > 5) vids[id] = { videoId: vid, publishDate: pd[i][COL_PUBLISHING.PUBLISH_DATE - 1] };
  }

  const last = sh.getLastRow();
  if (last <= ECON_TABLE_TOP) { ui.alert("No videos in the economics table — Refresh first."); return; }

  let updated = 0, noData = 0;
  for (var r = ECON_TABLE_TOP + 1; r <= last; r++) {
    const id = String(sh.getRange(r, 1).getValue()).trim();
    const v = vids[id];
    if (!v) continue;
    try {
      const start = v.publishDate instanceof Date
        ? Utilities.formatDate(v.publishDate, "UTC", "yyyy-MM-dd") : "2020-01-01";
      const end = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");
      const resp = YouTubeAnalytics.Reports.query({
        ids: "channel==MINE", startDate: start, endDate: end,
        metrics: "views,estimatedRevenue",
        filters: "video==" + v.videoId
      });
      if (resp && resp.rows && resp.rows.length) {
        sh.getRange(r, 9).setValue(resp.rows[0][0] || 0);   // Views
        sh.getRange(r, 10).setValue(resp.rows[0][1] || 0);  // Revenue $
        updated++;
      } else { noData++; }
    } catch (e) {
      Logger.log("Economics revenue " + id + ": " + e.message);
      noData++;
    }
    Utilities.sleep(200);
  }
  SpreadsheetApp.flush();
  ui.alert("📈 Revenue pulled",
    "Updated: " + updated + " video(s)\nNo data / not monetized: " + noData +
    "\n\nNET, ROI and Status recompute automatically.", ui.ButtonSet.OK);
}

/* NOTE — for EXACT Claude cost (not an estimate), log usage per video: have
   callClaude() append {contentId, stage, input_tokens, output_tokens, cache_read}
   to a hidden Usage_Log tab (it already reads json.usage), then sum by contentId
   here instead of estimating from output size. Requires passing contentId into
   callClaude at each call site. */
