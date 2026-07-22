/* ============================================================================
   research_bridge.gs — GovernX
   Bridge between the Node verified-claims engine and the Google Sheet.

   The engine (governx-remotion/src/server/research) does the hard work:
   fetch → parse → quote-gate → number-gate → hedge-gate → adversarial verify
   → attribution → relevance tiering → conflict detection.

   This file only orchestrates and presents. It never invents a claim.

   WHY ASYNC: Apps Script runs on Google's servers, so it cannot reach
   localhost. It needs the ngrok HTTPS URL. And UrlFetchApp times out long
   before an uncached research build finishes, so we POST a job, get an id,
   and poll until it's done.

   SETUP (once):
     Project Settings ⚙ → Script Properties → add
        RESEARCH_BASE_URL = https://<your-ngrok-subdomain>.ngrok-free.app
     Then run: setupResearchBridge()
     Then add this line inside your existing onOpen():  buildResearchMenu();

   SAFETY: additive. Own tabs, own constants. Touches nothing existing.
   ============================================================================ */

const SHEET_RB = {
  SOURCES  : "Research_Sources",
  CLAIMS   : "Research_Claims",
  CONFLICTS: "Research_Conflicts",
  DM       : "Research_Data_Moments"
};

const RB_ARCHIVE_FOLDER = "GovernX Research Archive";

// Column orders — index-safe, referenced by header position below.
// Claim_Type is appended LAST on purpose: existing index-based readers below
// keep working, and it feeds EVIDENCE_TYPE when publishing to Research Database.
const COL_RB_CLAIMS = ["Use", "Claim_ID", "Status", "Tier", "Relevance", "Attribution",
  "Value", "Statement", "Quote", "Source_Title", "Source_URL",
  "Verdict", "Overstated", "Reason", "Conflict_With", "Conflict_Note", "Claim_Type"];

const COL_RB_SOURCES = ["Source_#", "Title", "Publisher", "Type", "URL",
  "Words", "Hash", "Drive_Archive", "OK", "Error"];

const COL_RB_CONFLICTS = ["Metric", "Unit", "Role", "Claim_IDs", "Values", "How_to_present"];

const COL_RB_DM = ["DM_ID", "Claim_ID", "Value", "Label", "Source_Label",
  "Attribution", "Recommended_Visual"];

/* Research inputs live ON the idea, not in a prompt.
   Appended after PIPELINE_STATUS (col 14) so config.gs's COL_IDEA is untouched.
   Why this matters: the engine's cache is keyed on document + brief + model.
   A brief retyped with different punctuation is a cache miss AND a different
   set of claims from the same documents. Storing it makes research reproducible
   per idea instead of dependent on what you happened to type that day. */
const COL_RB_IDEA = { BRIEF: 15, URLS: 16, EDGAR: 17 };
const RB_IDEA_HEADERS = ["Research_Brief", "Source_URLs", "EDGAR_Query"];


// ── Menu ─────────────────────────────────────────────────────────────────────
// Call this from inside your existing onOpen().
function buildResearchMenu() {
  SpreadsheetApp.getUi().createMenu("🔎 Research")
    .addItem("① Setup research tabs",            "setupResearchBridge")
    .addItem("①b Add research columns to Idea Catalogue", "addResearchColumnsToIdeaCatalogue")
    .addSeparator()
    .addItem("② Run verified research…",         "runVerifiedResearch")
    .addItem("③ Approve ticked claims → Data Moments", "approveSelectedClaims")
    .addItem("④ Validate before render (blocks)", "validateBeforeRender")
    .addSeparator()
    .addItem("⑤ Publish → Research Database (feeds Stage 3)", "publishToResearchDatabase")
    .addItem("⑥ Validate scene numbers (run after Stage 4B)", "validateSceneNumbers")
    .addSeparator()
    .addItem("Check engine is reachable",        "pingResearchEngine")
    .addToUi();
}


// ── Setup ────────────────────────────────────────────────────────────────────
function setupResearchBridge() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  rbSheet_(ss, SHEET_RB.SOURCES,   COL_RB_SOURCES);
  rbSheet_(ss, SHEET_RB.CONFLICTS, COL_RB_CONFLICTS);
  rbSheet_(ss, SHEET_RB.DM,        COL_RB_DM);

  const claims = rbSheet_(ss, SHEET_RB.CLAIMS, COL_RB_CLAIMS);
  // "Use" is a checkbox you tick to promote a claim into the video.
  claims.getRange(2, 1, Math.max(claims.getMaxRows() - 1, 1), 1).insertCheckboxes();

  // Never let a column problem abort tab creation — report it instead.
  let colMsg = "";
  try {
    colMsg = rbEnsureIdeaColumns_(ss)
      ? "\n\nAdded to Idea Catalogue: " + RB_IDEA_HEADERS.join(", ")
      : "\n\nIdea Catalogue columns already present.";
  } catch (e) {
    colMsg = "\n\n⚠ Could not add Idea Catalogue columns: " + e.message +
             "\nRun 'Add research columns to Idea Catalogue' separately.";
  }

  SpreadsheetApp.getUi().alert("✅ Research tabs ready",
    "Created: " + Object.values(SHEET_RB).join(", ") + colMsg +
    "\n\nNext: Script Properties → RESEARCH_BASE_URL = your ngrok https URL." +
    "\nThen run 'Check engine is reachable'.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// Add Research_Brief / Source_URLs / EDGAR_Query to Idea Catalogue if absent.
// Idempotent — safe to run repeatedly.
function rbEnsureIdeaColumns_(ss) {
  const sh = ss.getSheetByName(typeof SHEET !== "undefined" ? SHEET.IDEA : "Idea Catalogue");
  if (!sh) return false;

  // Grow the sheet FIRST. getRange() on a column that doesn't exist yet throws,
  // so the existence check has to come after the columns exist.
  const missing = COL_RB_IDEA.EDGAR - sh.getMaxColumns();
  if (missing > 0) sh.insertColumnsAfter(sh.getMaxColumns(), missing);

  const first = String(sh.getRange(1, COL_RB_IDEA.BRIEF).getValue()).trim();
  if (first === RB_IDEA_HEADERS[0]) return false;   // already labelled

  sh.getRange(1, COL_RB_IDEA.BRIEF, 1, 3).setValues([RB_IDEA_HEADERS])
    .setFontWeight("bold").setBackground("#1F3A5F").setFontColor("#FFFFFF");
  sh.setColumnWidth(COL_RB_IDEA.BRIEF, 320);
  sh.setColumnWidth(COL_RB_IDEA.URLS,  360);
  sh.setColumnWidth(COL_RB_IDEA.EDGAR, 240);
  return true;
}

// Runnable on its own, so a failure here never leaves you re-running tab setup.
function addResearchColumnsToIdeaCatalogue() {
  const ui = SpreadsheetApp.getUi();
  const added = rbEnsureIdeaColumns_(SpreadsheetApp.getActiveSpreadsheet());
  ui.alert(added ? "✅ Columns added" : "✅ Already present",
    RB_IDEA_HEADERS.join(", ") + "\n\nIdea Catalogue columns O, P, Q.", ui.ButtonSet.OK);
}

// getActiveIdeaRow() doesn't hand back the row index, and we need it to read
// and write the three research-input cells.
function rbActiveIdea_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const ideaName = (typeof SHEET !== "undefined") ? SHEET.IDEA : "Idea Catalogue";
  if (sh.getName() !== ideaName) {
    SpreadsheetApp.getUi().alert("Select a row in the " + ideaName + " tab first.");
    return null;
  }
  const row = sh.getActiveCell().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert("Select a data row, not the header."); return null; }

  const id = String(sh.getRange(row, 1).getValue()).trim();
  if (!id) { SpreadsheetApp.getUi().alert("This row has no ID yet."); return null; }

  return { row, sheet: sh, id, company: sh.getRange(row, 2).getValue() };
}

function rbSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#1F3A5F").setFontColor("#FFFFFF");
  sh.setFrozenRows(1);
  sh.setTabColor("#1F3A5F");
  return sh;
}


// ── Engine plumbing ──────────────────────────────────────────────────────────
function rbBaseUrl_() {
  const url = PropertiesService.getScriptProperties().getProperty("RESEARCH_BASE_URL");
  if (!url) throw new Error(
    "RESEARCH_BASE_URL is not set.\n\nProject Settings ⚙ → Script Properties → add:\n" +
    "RESEARCH_BASE_URL = https://<your-ngrok>.ngrok-free.app");
  return url.replace(/\/+$/, "");
}

function rbFetch_(path, options) {
  const opts = options || {};
  opts.muteHttpExceptions = true;
  opts.headers = Object.assign({ "ngrok-skip-browser-warning": "1" }, opts.headers || {});
  const resp = UrlFetchApp.fetch(rbBaseUrl_() + path, opts);
  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) throw new Error("Engine HTTP " + code + ": " + body.slice(0, 300));
  return JSON.parse(body);
}

function pingResearchEngine() {
  try {
    const health = rbFetch_("/health", { method: "get" });
    SpreadsheetApp.getUi().alert("✅ Engine reachable", JSON.stringify(health), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("❌ Cannot reach engine", e.message +
      "\n\nIs `npm start` running, and is ngrok pointed at port 3000?", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// ── ② Run verified research ──────────────────────────────────────────────────
function runVerifiedResearch() {
  const ui = SpreadsheetApp.getUi();

  // Same contract as every other stage: act on the selected Idea Catalogue row.
  const idea = rbActiveIdea_();
  if (!idea) return;
  rbEnsureIdeaColumns_(SpreadsheetApp.getActiveSpreadsheet());

  const company = idea.company;
  let brief   = String(idea.sheet.getRange(idea.row, COL_RB_IDEA.BRIEF).getValue()).trim();
  let urlsRaw = String(idea.sheet.getRange(idea.row, COL_RB_IDEA.URLS ).getValue()).trim();
  let edgarQ  = String(idea.sheet.getRange(idea.row, COL_RB_IDEA.EDGAR).getValue()).trim();

  // ── No typing required ─────────────────────────────────────────────────────
  // This used to open three prompts on the first run. Everything they asked for
  // already exists in the sheet by the time ② runs, so derive it instead:
  //   brief  ← Master Content Table (Stage 1)
  //   urls   ← the real http links Stage 2 put in the Research Database
  //            (Stage 2 writes "Search: …" when it cannot confirm a URL — those
  //             are deliberately NOT passed to the engine, there is nothing to fetch)
  //   edgar  ← left blank; it only helps for US filers and is rarely worth the call
  // Anything already present in cols O/P/Q wins, so a manual override still works.
  var derived = [];
  if (!brief)   { brief   = rbAutoBrief_(idea);          if (brief)   derived.push("brief"); }
  if (!urlsRaw) { urlsRaw = rbAutoUrls_(idea.id);        if (urlsRaw) derived.push("source URLs"); }

  if (derived.length) {
    idea.sheet.getRange(idea.row, COL_RB_IDEA.BRIEF).setValue(brief);
    idea.sheet.getRange(idea.row, COL_RB_IDEA.URLS ).setValue(urlsRaw);
    idea.sheet.getRange(idea.row, COL_RB_IDEA.EDGAR).setValue(edgarQ);
    SpreadsheetApp.flush();
    Logger.log("Research ②: auto-filled " + derived.join(" + ") + " for " + idea.id);
  }

  const payload = {
    company: company,
    brief  : brief,
    urls   : urlsRaw.split(",").map(function (s) { return s.trim(); }).filter(String),
    edgarQueries: edgarQ ? [{ q: edgarQ, limit: 2 }] : []
  };
  if (!payload.urls.length && !payload.edgarQueries.length) {
    // Stage 2 writes "Search: …" instead of inventing a URL it cannot confirm, so
    // a topic can legitimately reach ② with nothing fetchable. Say exactly that,
    // and show the searches Stage 2 suggested so the gap is easy to close.
    const hints = rbSearchHints_(idea.id);
    ui.alert("No fetchable sources yet — " + idea.id,
      "Nothing was auto-filled because Stage 2 could not confirm a single URL for " +
      company + ".\n\n" +
      (hints.length
        ? "Stage 2 suggested these searches:\n\n  • " + hints.slice(0, 6).join("\n  • ") +
          "\n\nRun one or two, then paste the resulting document links into\n" +
          "Idea Catalogue → Source_URLs (col P), comma-separated, and re-run ②."
        : "Paste at least one document link into Idea Catalogue → Source_URLs (col P) and re-run ②.") +
      "\n\nTip: sec.gov blocks automated fetching — download the PDF and host it\n" +
      "on Drive (anyone-with-link), then use that link instead.",
      ui.ButtonSet.OK);
    return;
  }

  // Start the job.
  const started = rbFetch_("/research/job", {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload)
  });
  const jobId = started.jobId;

  // Poll. Cached builds land in ~40s; a cold build can take 3–4 minutes.
  let data = null;
  for (let i = 0; i < 22; i++) {          // 22 × 15s ≈ 5.5 min, inside the 6-min cap
    Utilities.sleep(15000);
    const poll = rbFetch_("/research/job/" + jobId, { method: "get" });
    if (poll.status === "done")  { data = poll; break; }
    if (poll.status === "error") throw new Error("Research failed: " + poll.error);
  }
  if (!data) throw new Error("Research timed out after ~5.5 min. Check the server window.");

  writeResearchResults_(data, idea.id);

  const s = data.stats;
  ui.alert("🔎 Research complete",
    "Sources: " + s.sourcesOk + "/" + s.sources +
    "\nClaims: " + s.claims + "   Verified: " + s.verified + " (" + s.verifiedRate + "%)" +
    "\nUsable (verified, not boilerplate): " + s.usable +
    "\nCore: " + s.core + "   Needs review: " + s.needsReview + "   Rejected: " + s.rejected +
    "\nConflicts flagged: " + s.conflicts +
    ((data.warnings && data.warnings.length) ? "\n\n⚠ " + data.warnings.join("\n⚠ ") : "") +
    "\n\nNow tick the 'Use' box on the claims you want, then run step ③.",
    ui.ButtonSet.OK);
}

function rbPrompt_(ui, title, help) {
  const r = ui.prompt(title, help, ui.ButtonSet.OK_CANCEL);
  return r.getSelectedButton() === ui.Button.OK ? r.getResponseText().trim() : null;
}


/* ── ②b SUGGEST CLAIMS ─────────────────────────────────────────────────────────
   A run returns 40-plus claims and ticking them all is the wrong instinct: the
   Nissan set had 14 claims sitting in conflict groups — C1/C3/C9 are the SAME
   $140M figure in different wordings — so approving everything publishes one
   metric four times and the film repeats itself.

   This ticks a defensible subset and unticks the rest. It only ever uses signals
   the engine itself produced; it does not judge the content:
     • never tick anything that is not Verified, or flagged Overstated
     • within a conflict group keep exactly ONE claim — the highest scoring
     • prefer Core over Supporting, a hard number over a qualitative statement,
       and regulator/court attribution over company or press
     • cap the total, because a 5-minute film cannot carry 40 figures
   Everything stays reviewable: the checkboxes are the output, and you can
   re-tick anything by hand before ③.
   ---------------------------------------------------------------------------- */
/* How many figures a film can actually carry, derived from the SAME scene
   targets Stage 4 uses (DIRECTOR_SPEC §3) so the two never drift:

     format                scenes    data scenes   figures
     Short (< 90s)          8–12         ~5           10
     Standard (4–7 min)    26–30        ~14           24
     Deep Dive (10–15 min) 45–60        ~28           45

   Roughly half the scenes carry data, and a data scene uses 1–4 figures (a KPI
   dashboard or data wall eats 3–4 on its own), so figures ≈ scenes × 0.85.
   A fixed 18 was too tight for a Standard film and far too tight for a Deep Dive. */
const RB_CLAIM_CAP_BY_FORMAT = {
  "Short (< 90s)"        : 10,
  "Standard (4-7 min)"   : 24,
  "Standard (4–7 min)"   : 24,   // en-dash variant used in some rows
  "Deep Dive (10-15 min)": 45,
  "Deep Dive (10–15 min)": 45
};
const RB_CLAIM_CAP_DEFAULT = 24;

function rbRecommendedCap_() {
  try {
    const idea = rbActiveIdea_();
    if (!idea) return { cap: RB_CLAIM_CAP_DEFAULT, format: "(no row selected)" };
    const fmt = String(idea.sheet.getRange(idea.row, COL_IDEA.TARGET_FORMAT).getValue() || "").trim();
    return { cap: RB_CLAIM_CAP_BY_FORMAT[fmt] || RB_CLAIM_CAP_DEFAULT, format: fmt || "(not set)" };
  } catch (e) {
    return { cap: RB_CLAIM_CAP_DEFAULT, format: "(unknown)" };
  }
}

function suggestClaims() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_RB.CLAIMS);
  if (!sh || sh.getLastRow() < 2) {
    ui.alert("No claims yet — run ② Run verified research first."); return;
  }

  const n = sh.getLastRow() - 1;
  const rows = sh.getRange(2, 1, n, COL_RB_CLAIMS.length).getValues();

  const C = { USE: 0, ID: 1, STATUS: 2, TIER: 3, REL: 4, ATTR: 5, VALUE: 6,
              STMT: 7, OVER: 12, CONFLICT: 14 };

  const isNum = (v) => /\d/.test(String(v || ""));
  const yes   = (v) => { const s = String(v || "").trim().toLowerCase(); return s && s !== "no" && s !== "false"; };

  // ── score ──────────────────────────────────────────────────────────────────
  const claim = rows.map(function (r, i) {
    const status = String(r[C.STATUS] || "").trim();
    const tier   = String(r[C.TIER]   || "").trim();
    const rel    = String(r[C.REL]    || "").trim().toLowerCase();
    const attr   = String(r[C.ATTR]   || "").trim().toLowerCase();

    let score = 0, why = [];
    if (/^core$/i.test(tier))            { score += 3; why.push("Core"); }
    else if (/^supporting$/i.test(tier)) { score += 1; why.push("Supporting"); }
    if (isNum(r[C.VALUE]))               { score += 3; why.push("has figure"); }
    if (rel === "high")                  { score += 2; }
    else if (rel === "medium")           { score += 1; }
    if (/regulat|court/.test(attr))      { score += 2; why.push("official source"); }
    else if (/compan|self/.test(attr))   { score += 1; }

    const blocked = status !== "Verified" ? "not Verified (" + status + ")"
                  : yes(r[C.OVER])       ? "flagged Overstated"
                  : /^boilerplate$/i.test(tier) ? "boilerplate"
                  : null;

    return { i: i, id: String(r[C.ID] || "").trim(), score: score, why: why,
             blocked: blocked, conflict: String(r[C.CONFLICT] || "").trim(),
             stmt: String(r[C.STMT] || "").trim() };
  });

  const byId = {};
  claim.forEach(function (c) { byId[c.id] = c; });

  // ── one claim per conflict group ───────────────────────────────────────────
  // Conflict_With lists the other members, so walk the links to build each group.
  const groupOf = {}, groups = [];
  claim.forEach(function (c) {
    if (c.blocked || !c.conflict || groupOf[c.id] !== undefined) return;
    const stack = [c.id], members = [];
    while (stack.length) {
      const id = stack.pop();
      if (groupOf[id] !== undefined || !byId[id] || byId[id].blocked) continue;
      groupOf[id] = groups.length;
      members.push(id);
      String(byId[id].conflict).split(",").forEach(function (x) {
        const t = x.trim(); if (t && groupOf[t] === undefined) stack.push(t);
      });
    }
    if (members.length) groups.push(members);
  });

  const droppedForConflict = {};
  groups.forEach(function (members) {
    let best = null;
    members.forEach(function (id) {
      const c = byId[id];
      if (!best || c.score > best.score) best = c;
    });
    members.forEach(function (id) {
      if (id !== best.id) droppedForConflict[id] = best.id;
    });
  });

  // ── choose ─────────────────────────────────────────────────────────────────
  const eligible = claim
    .filter(function (c) { return !c.blocked && !droppedForConflict[c.id]; })
    .sort(function (a, b) { return b.score - a.score; });

  // The cap follows the video's Target Format, but it is a judgement call — so
  // recommend and let the user override. Empty input accepts the recommendation.
  const rec = rbRecommendedCap_();
  const suggested = Math.min(rec.cap, eligible.length);
  const ans = ui.prompt(
    "How many claims should this film carry?",
    "Format: " + rec.format + "\n" +
    "Recommended: " + suggested + "   (a " + rec.format + " film carries about " + rec.cap + " figures)\n\n" +
    eligible.length + " claims are usable after removing unverified and duplicate figures.\n\n" +
    "Press OK to accept " + suggested + ", or type a different number:",
    ui.ButtonSet.OK_CANCEL
  );
  if (ans.getSelectedButton() !== ui.Button.OK) return;
  const typed = parseInt(String(ans.getResponseText()).trim(), 10);
  const cap = (Number.isFinite(typed) && typed > 0) ? typed : suggested;

  const keep = {};
  eligible.slice(0, cap).forEach(function (c) { keep[c.id] = true; });

  // ── write the checkboxes ───────────────────────────────────────────────────
  const use = claim.map(function (c) { return [ !!keep[c.id] ]; });
  sh.getRange(2, 1, n, 1).setValues(use);
  SpreadsheetApp.flush();

  // ── report ─────────────────────────────────────────────────────────────────
  const blockedList  = claim.filter(function (c) { return c.blocked; });
  const overCap      = eligible.length - Math.min(eligible.length, cap);
  const conflictKeys = Object.keys(droppedForConflict);

  let msg = "Ticked " + Object.keys(keep).length + " of " + n + " claims" +
            "  (cap " + cap + " — " + rec.format + ").\n\n";
  if (blockedList.length) {
    msg += "EXCLUDED — not usable (" + blockedList.length + "):\n";
    blockedList.slice(0, 6).forEach(function (c) { msg += "  " + c.id + " — " + c.blocked + "\n"; });
    if (blockedList.length > 6) msg += "  …and " + (blockedList.length - 6) + " more\n";
    msg += "\n";
  }
  if (conflictKeys.length) {
    msg += "DE-DUPLICATED — same metric, kept the strongest (" + conflictKeys.length + "):\n";
    conflictKeys.slice(0, 6).forEach(function (id) { msg += "  " + id + " → kept " + droppedForConflict[id] + "\n"; });
    if (conflictKeys.length > 6) msg += "  …and " + (conflictKeys.length - 6) + " more\n";
    msg += "\n";
  }
  if (overCap > 0) msg += "TRIMMED — " + overCap + " more were usable but sit below the cap of " + cap + ".\n" +
                          "   Re-run ②b and enter a higher number if you want them.\n\n";
  msg += "Review the Use column and re-tick anything you want, then run ③.";

  ui.alert("②b Claim suggestions — " + SHEET_RB.CLAIMS, msg, ui.ButtonSet.OK);
}


/* ── Auto-fill for ② (replaces the three prompts) ──────────────────────────────
   All three inputs already exist in the sheet by the time ② runs. Deriving them
   also makes the brief byte-stable across re-runs, which keeps the engine's cache
   key (document + brief + model) hitting instead of rebuilding.
   ---------------------------------------------------------------------------- */

// Brief ← Master Content Table (Stage 1). Falls back to the idea row's angle.
function rbAutoBrief_(idea) {
  var m = {};
  try { m = getMasterContent(idea.id) || {}; } catch (e) {}
  const bits = [];
  const angle = String(m.primaryAngle || idea.initialAngle || "").trim();
  const insight = String(m.coreInsight || "").trim();
  const disc = String(m.discipline || "").trim();
  const cps = String(m.checkpoints || "").trim();

  bits.push("Verify the governance failure at " + String(idea.company || "").trim() + ".");
  if (angle)   bits.push("Focus: " + angle);
  if (insight) bits.push("Core insight to test: " + insight);
  if (cps)     bits.push("Key moments: " + cps.replace(/\s*(->|→)\s*/g, " → "));
  bits.push("Extract specific verifiable figures (amounts, counts, percentages), the dates " +
            "they attach to, who is asserting each one, and any regulator or court finding.");
  if (disc)    bits.push("Discipline lens: " + disc + ".");
  bits.push("Prioritise primary and official sources over commentary.");
  return bits.join(" ");
}

// Source URLs ← the real links Stage 2 wrote into the Research Database.
// "Search: …" rows are skipped: there is no document behind them to fetch.
function rbAutoUrls_(contentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(rbResearchSheetName_());
    if (!sh || sh.getLastRow() < 2) return "";
    const C = rbResearchCols_();
    const data = sh.getDataRange().getValues();
    const seen = {}, out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][C.ID - 1]).trim() !== contentId) continue;
      const link = String(data[i][C.SOURCE_LINK - 1] || "").trim();
      if (!/^https?:\/\//i.test(link)) continue;      // skips "Search: …"
      if (seen[link]) continue;
      seen[link] = true;
      out.push(link);
      if (out.length >= 8) break;                     // engine budget
    }
    return out.join(", ");
  } catch (e) { return ""; }
}

// The "Search: …" queries Stage 2 left behind — shown when there is nothing to fetch.
function rbSearchHints_(contentId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(rbResearchSheetName_());
    if (!sh || sh.getLastRow() < 2) return [];
    const C = rbResearchCols_();
    const data = sh.getDataRange().getValues();
    const out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][C.ID - 1]).trim() !== contentId) continue;
      const link = String(data[i][C.SOURCE_LINK - 1] || "").trim();
      if (/^search\s*:/i.test(link)) out.push(link.replace(/^search\s*:\s*/i, ""));
    }
    return out;
  } catch (e) { return []; }
}


// ── Write results ────────────────────────────────────────────────────────────
function writeResearchResults_(data, contentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Sources (+ archive full text to Drive so the evidence survives the session)
  const srcSh = rbSheet_(ss, SHEET_RB.SOURCES, COL_RB_SOURCES);
  rbClearBody_(srcSh);
  const folder = rbArchiveFolder_(contentId);
  const srcRows = (data.sources || []).map(function (s) {
    let link = "";
    if (s.ok && s.text) {
      const f = folder.createFile(s.hash + "__" + (s.title || "source").replace(/[^\w]+/g, "_").slice(0, 60) + ".txt",
        "URL: " + s.finalUrl + "\nTITLE: " + s.title + "\nSHA256(16): " + s.hash + "\n\n" + s.text);
      link = f.getUrl();
    }
    return [s.sourceIndex, s.title, s.publisher, s.sourceType, s.url,
            s.wordCount, s.hash, link, s.ok ? "Yes" : "No", s.error || ""];
  });
  if (srcRows.length) srcSh.getRange(2, 1, srcRows.length, COL_RB_SOURCES.length).setValues(srcRows);

  // Claims — the human review queue.
  const clSh = rbSheet_(ss, SHEET_RB.CLAIMS, COL_RB_CLAIMS);
  rbClearBody_(clSh);
  const clRows = (data.claims || []).map(function (c) {
    return [false, c.claimId, c.status, c.tier, c.relevance, c.attribution,
            c.value, c.statement, c.quote, c.sourceTitle, c.sourceUrl,
            c.verdict, c.overstatement ? "YES" : "", c.reason,
            (c.conflictWith || []).join(", "), c.conflictNote || "", c.claimType || ""];
  });
  if (clRows.length) {
    // Value/Statement/Quote must be PLAIN TEXT before the write. A claim whose
    // value is a date ("November 19, 2018") is otherwise coerced by Sheets into a
    // date serial — C18 came back as 43423 and C19 as 43563, which would have
    // rendered as meaningless digits on screen. Formatting the columns as "@"
    // first stops the coercion at the source.
    const VALUE_COL = 7, STMT_COL = 8, QUOTE_COL = 9;
    clSh.getRange(2, VALUE_COL, clRows.length, 3).setNumberFormat("@");
    clSh.getRange(2, 1, clRows.length, COL_RB_CLAIMS.length).setValues(clRows);
    clSh.getRange(2, 1, clRows.length, 1).insertCheckboxes();
    rbColorClaims_(clSh, data.claims);
  }

  // Conflicts.
  const cfSh = rbSheet_(ss, SHEET_RB.CONFLICTS, COL_RB_CONFLICTS);
  rbClearBody_(cfSh);
  const cfRows = (data.conflicts || []).map(function (g) {
    const ids = g.claimIndexes.map(function (i) { return data.claims[i].claimId; });
    const vals = g.claimIndexes.map(function (i) { return data.claims[i].value; });
    return [g.metric, g.unit || "", g.role || "", ids.join(", "), vals.join("  |  "), g.note];
  });
  if (cfRows.length) cfSh.getRange(2, 1, cfRows.length, COL_RB_CONFLICTS.length).setValues(cfRows);
}

// Green = safe to use. Amber = needs a human decision. Red = do not use.
function rbColorClaims_(sh, claims) {
  claims.forEach(function (c, i) {
    const row = sh.getRange(i + 2, 1, 1, COL_RB_CLAIMS.length);
    if (c.status === "Rejected")            row.setBackground("#FCE8E6");
    else if (c.status === "Needs Review")   row.setBackground("#FEF7E0");
    else if (c.tier === "Boilerplate")      row.setBackground("#F1F3F4");
    else if (c.tier === "Core")             row.setBackground("#E6F4EA");
  });
}

function rbClearBody_(sh) {
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clear();
}

// Archive root, then ONE SUBFOLDER PER CONTENT ID. The archive used to be flat:
// every topic's sources landed together, named by hash, so there was no way to
// tell which document belonged to which film. Fine at 3 topics, unusable at 50.
// Mirrors how getOrCreateContentFolder organises production packages.
function rbArchiveFolder_(contentId) {
  const it = DriveApp.getFoldersByName(RB_ARCHIVE_FOLDER);
  const root = it.hasNext() ? it.next() : DriveApp.createFolder(RB_ARCHIVE_FOLDER);
  const id = String(contentId || "").trim();
  if (!id) return root;                       // no id known — keep old behaviour
  const sub = root.getFoldersByName(id);
  return sub.hasNext() ? sub.next() : root.createFolder(id);
}


// ── ③ Approve ticked claims → Data Moments ───────────────────────────────────
function approveSelectedClaims() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clSh = ss.getSheetByName(SHEET_RB.CLAIMS);
  if (!clSh || clSh.getLastRow() < 2) { ui.alert("No claims. Run step ② first."); return; }

  const rows = clSh.getRange(2, 1, clSh.getLastRow() - 1, COL_RB_CLAIMS.length).getValues();
  const picked = [], blocked = [];

  rows.forEach(function (r) {
    if (r[0] !== true) return;                       // Use checkbox not ticked
    const claimId = r[1], status = r[2], tier = r[3], attribution = r[5], value = r[6];
    if (status !== "Verified")   { blocked.push(claimId + " — status is " + status); return; }
    if (tier === "Boilerplate")  { blocked.push(claimId + " — boilerplate"); return; }
    const val = rbCellText_(value);                     // repairs a coerced date
    picked.push([ "DM_" + (picked.length + 1), claimId, val,
                  rbShortLabel_(r[7]),                                // Label ← short caption
                  "Source: " + r[9],                                  // Source_Label
                  attribution,
                  rbRecommendVisual_(val, r[16], r[8], attribution) ]);
  });

  if (blocked.length) {
    ui.alert("🚫 Some ticked claims cannot be used",
      blocked.join("\n") + "\n\nUntick them, or resolve them first.", ui.ButtonSet.OK);
    return;
  }
  if (!picked.length) { ui.alert("Tick the 'Use' box on the claims you want first."); return; }

  const dmSh = rbSheet_(ss, SHEET_RB.DM, COL_RB_DM);
  rbClearBody_(dmSh);
  dmSh.getRange(2, 3, picked.length, 1).setNumberFormat("@");   // Value stays text
  dmSh.getRange(2, 1, picked.length, COL_RB_DM.length).setValues(picked);

  const visuals = {};
  picked.forEach(function (p) { visuals[p[6]] = (visuals[p[6]] || 0) + 1; });
  const summary = Object.keys(visuals).sort(function (a, b) { return visuals[b] - visuals[a]; })
    .map(function (k) { return "  " + String(visuals[k]).padStart(2) + "  " + k; }).join("\n");

  ui.alert("✅ " + picked.length + " data moment(s) created",
    "Recommended_Visual has been filled from each figure's shape:\n\n" + summary +
    "\n\nThese are suggestions for Stage 4B — override any of them by editing the\n" +
    "column. Run step ④ before rendering.", ui.ButtonSet.OK);
}


/* ── Data-moment helpers ──────────────────────────────────────────────────────
   Three things the Data Moments tab was not doing, all of which cost the director
   information it needed:
     • Value could be a date SERIAL (43423) instead of a date
     • Label was the first 90 chars of a full sentence — unusable as an on-screen
       caption, which is why scene captions had to be written by hand
     • Recommended_Visual was left empty on every row
   ---------------------------------------------------------------------------- */

// A value Sheets coerced into a Date comes back as a Date object; render it the
// way a viewer reads it rather than as a serial number.
function rbCellText_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || "UTC", "d MMMM yyyy");
  }
  const s = String(v == null ? "" : v).trim();
  // A bare 5-digit number in the date-serial range is almost certainly a coerced
  // date that was stored before the "@" formatting fix.
  if (/^\d{5}$/.test(s)) {
    const n = parseInt(s, 10);
    if (n > 25000 && n < 60000) {                       // ~1968 → ~2064
      const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return Utilities.formatDate(d, "UTC", "d MMMM yyyy");
    }
  }
  return s;
}

// A caption, not a sentence. Take the first clause and cap it — the full
// statement stays available on the claim row.
function rbShortLabel_(statement) {
  let s = String(statement || "").trim().replace(/\s+/g, " ");
  s = s.split(/(?:\.\s|;\s|\s—\s|,\s(?=which|and that))/)[0];   // first clause
  if (s.length > 64) {
    const cut = s.slice(0, 64);
    s = cut.slice(0, Math.max(cut.lastIndexOf(" "), 40)) + "…";
  }
  return s.replace(/[.,;:]$/, "");
}

// Suggest the component that suits the SHAPE of the figure. Every output is a
// real REMOTION_DATA type the Stage 4B director already knows.
function rbRecommendVisual_(value, claimType, quote, attribution) {
  const v = String(value || "").trim();
  const q = String(quote || "").trim();
  const t = String(claimType || "").toLowerCase();

  // EVIDENCE_CARD is the hero scene — the verbatim trust unit — and a film wants
  // ONE, maybe two. An earlier rule ("long quote + regulator attribution") fired
  // on 17 of 40 claims, because every SEC claim has both. It is now reserved for
  // a finding with NO figure: if there is a number, the number leads and the
  // right component follows the number's shape.
  if (!/\d/.test(v)) return q.length >= 40 ? "EVIDENCE_CARD" : "STAT_POSTER";

  if (/%|percent/i.test(v))                            return "PROGRESS_GAUGE";
  if (/^(19|20)\d{2}$/.test(v) ||
      (/\b(19|20)\d{2}\b/.test(v) &&
       /january|february|march|april|may|june|july|august|september|october|november|december/i.test(v)))
                                                       return "CHECKPOINT";
  if (/billion|bn\b|[\d.]+\s*b\b/i.test(v))            return "COUNTER_ANIMATION";
  if (/compar|versus|vs\b|before|after/.test(t))       return "SPLIT_COMPARISON";
  return "DATA_CALLOUT";
}


/* ============================================================================
   ⑤ PUBLISH TO `Research Database` — the link into the main pipeline.

   This is the whole point. Stages 3–11 never read Research_Claims; they read
   the `Research Database` tab, keyed by Content ID, taking only the rows where
   USED_IN_SCRIPT = "YES". Publishing writes verified claims into that exact
   11-column schema, so Stage 3 writes its script from VERBATIM QUOTES with
   source links instead of from model recall.

   Nothing in pipeline.gs is modified. The original generateResearchDatabase()
   stays as a fallback. This is a parallel, additive Stage 2.

   Row policy:
     Verified + not Boilerplate  → USED_IN_SCRIPT = "YES"   (script may use it)
     Needs Review                → USED_IN_SCRIPT = ""       (visible, excluded)
     Rejected / Boilerplate      → USED_IN_SCRIPT = "NO"     (documented, barred)
   ============================================================================ */

// Fall back to local constants if config.gs hasn't loaded (standalone testing).
function rbResearchSheetName_() {
  return (typeof SHEET !== "undefined" && SHEET.RESEARCH) ? SHEET.RESEARCH : "Research Database";
}
function rbResearchCols_() {
  return (typeof COL_RESEARCH !== "undefined") ? COL_RESEARCH : {
    ID: 1, TOPIC: 2, SOURCE_TYPE: 3, DETAILS: 4, SOURCE_LINK: 5, KEY_INSIGHT: 6,
    EVIDENCE_TYPE: 7, TIMESTAMP: 8, RELEVANCE: 9, USED_IN_SCRIPT: 10, NOTE: 11
  };
}

function publishToResearchDatabase() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const clSh = ss.getSheetByName(SHEET_RB.CLAIMS);
  if (!clSh || clSh.getLastRow() < 2) { ui.alert("No claims. Run step ② first."); return; }

  // Same contract as every other stage: act on the selected Idea Catalogue row.
  const idea = getActiveIdeaRow();
  if (!idea) return;                       // getActiveIdeaRow() alerts on its own
  const contentId = idea.id;
  const topic     = idea.company;

  // Publishing replaces prior rows for this ID — say so before doing it.
  const go = ui.alert("Publish verified claims to " + rbResearchSheetName_(),
    "Content ID: " + contentId + "\nTopic: " + topic +
    "\n\nAny existing " + rbResearchSheetName_() + " rows for this ID will be replaced.",
    ui.ButtonSet.OK_CANCEL);
  if (go !== ui.Button.OK) return;

  const rows = clSh.getRange(2, 1, clSh.getLastRow() - 1, COL_RB_CLAIMS.length).getValues();
  const C = rbResearchCols_();
  const now = new Date();

  const out = [], counts = { yes: 0, blank: 0, no: 0 };

  rows.forEach(function (r) {
    const claimId = r[1], status = r[2], tier = r[3], relevance = r[4], attribution = r[5];
    const value = r[6], statement = r[7], quote = r[8], sourceTitle = r[9], sourceUrl = r[10];
    const overstated = r[12], conflictWith = r[14], conflictNote = r[15], claimType = r[16];
    if (!claimId) return;

    let used = "";
    if (status === "Verified" && tier !== "Boilerplate") { used = "YES";  counts.yes++; }
    else if (status === "Rejected" || tier === "Boilerplate") { used = "NO"; counts.no++; }
    else { counts.blank++; }   // Needs Review — present, deliberately unused

    // NOTE carries every caveat the script writer must respect,
    // plus the tier detail that RELEVANCE's dropdown can't hold.
    const notes = [claimId];
    notes.push("Tier: " + tier + (relevance ? " (" + relevance + "/5)" : ""));
    if (attribution)  notes.push("Asserted by: " + attribution);
    if (overstated === "YES") notes.push("⚠ OVERSTATED — source hedges; do not assert flatly");
    if (status === "Needs Review") notes.push("⚠ NOT VERIFIED — excluded from script");
    if (conflictWith) notes.push("⚠ CONFLICTS WITH " + conflictWith + ". " + conflictNote);

    const row = new Array(11).fill("");
    row[C.ID - 1]             = contentId;
    row[C.TOPIC - 1]          = topic;
    row[C.SOURCE_TYPE - 1]    = attribution || "Unclear";
    row[C.DETAILS - 1]        = quote;                    // ← the verbatim evidence
    row[C.SOURCE_LINK - 1]    = sourceUrl;
    row[C.KEY_INSIGHT - 1]    = statement + (value ? "  [" + value + "]" : "");
    row[C.EVIDENCE_TYPE - 1]  = claimType || "";
    row[C.TIMESTAMP - 1]      = now;
    row[C.RELEVANCE - 1]      = rbTierToRelevance_(tier);  // strict dropdown: High|Medium|Low
    row[C.USED_IN_SCRIPT - 1] = used;
    row[C.NOTE - 1]           = notes.join(" · ");
    out.push(row);
  });

  if (!out.length) { ui.alert("Nothing to publish."); return; }

  const rdSh = ss.getSheetByName(rbResearchSheetName_());
  if (!rdSh) { ui.alert("❌ '" + rbResearchSheetName_() + "' tab not found."); return; }

  // Idempotent: clear any rows this Content ID published before.
  rbDeleteRowsForContentId_(rdSh, contentId, C.ID);

  const startRow = rdSh.getLastRow() + 1;
  const target   = rdSh.getRange(startRow, 1, out.length, 11);

  // A strict data-validation rule anywhere in the target range makes setValues
  // throw and writes NOTHING. Retry once with validations cleared on our rows
  // rather than leaving a silent partial publish.
  try {
    target.setValues(out);
  } catch (e) {
    target.clearDataValidations();
    target.setValues(out);
    ui.alert("⚠ Data validation cleared",
      "Some columns rejected the written values, so validation was removed on " +
      "rows " + startRow + "–" + (startRow + out.length - 1) + ".\n\n" + e.message,
      ui.ButtonSet.OK);
  }

  // Colour by usability so the sheet reads at a glance.
  out.forEach(function (r, i) {
    const used = r[C.USED_IN_SCRIPT - 1];
    const bg = used === "YES" ? "#E6F4EA" : used === "NO" ? "#FCE8E6" : "#FEF7E0";
    rdSh.getRange(startRow + i, 1, 1, 11).setBackground(bg);
  });

  try { updatePipelineStatus_(contentId, "S2", "✅"); } catch (e) {}

  ui.alert("✅ Published to " + rbResearchSheetName_(),
    out.length + " row(s) written for " + contentId + ":\n\n" +
    "  USED_IN_SCRIPT = YES : " + counts.yes + "  (verified, script may cite)\n" +
    "  USED_IN_SCRIPT = ''  : " + counts.blank + "  (needs review — excluded)\n" +
    "  USED_IN_SCRIPT = NO  : " + counts.no + "  (rejected / boilerplate)\n\n" +
    "DETAILS holds the verbatim quote; SOURCE_LINK the document.\n\n" +
    "You can now run Stage 3 — Generate Script. It will read only the YES rows.",
    ui.ButtonSet.OK);
}

// `Research Database` column I is a strict dropdown: High | Medium | Low.
// The engine's tier is richer than that, so the tier + score live in NOTE and
// only the allowed token goes in the validated cell.
function rbTierToRelevance_(tier) {
  if (tier === "Core")        return "High";
  if (tier === "Supporting")  return "Medium";
  return "Low";                                  // Boilerplate / unknown
}

function rbDeleteRowsForContentId_(sh, contentId, idCol) {
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, idCol, sh.getLastRow() - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {                 // bottom-up: indices stay valid
    if (String(ids[i][0]).trim() === contentId) sh.deleteRow(i + 2);
  }
}


/* ============================================================================
   ⑥ VALIDATE SCENE NUMBERS — the guardrail's safety net.

   Stage 4 builds scenes partly from the pre-verification brief, so a fabricated
   figure can slip onto screen even after the prompt hardening. This scans every
   number in the Visual Library REMOTION_DATA / CHECKPOINT_EVENT for the selected
   idea and flags any that does not appear in the verified claim set.

   It reads the same `Research Database` rows the video was published from, so it
   needs no server call. Run it after Stage 4B, before Stage 8D.
   ============================================================================ */
function validateSceneNumbers() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idea = getActiveIdeaRow();
  if (!idea) return;

  // 1) Build the whitelist of verified number-tokens from Research Database.
  const rdName = rbResearchSheetName_(), C = rbResearchCols_();
  const rdSh = ss.getSheetByName(rdName);
  if (!rdSh) { ui.alert("No '" + rdName + "' tab. Publish verified research first (⑤)."); return; }

  const rd = rdSh.getDataRange().getValues();
  const verified = {};   // token → true
  let verifiedRows = 0;
  for (let i = 1; i < rd.length; i++) {
    if (String(rd[i][C.ID - 1]).trim() !== idea.id) continue;
    verifiedRows++;
    rbNumberTokens_(rd[i][C.DETAILS - 1] + " " + rd[i][C.KEY_INSIGHT - 1])
      .forEach(function (t) { verified[t] = true; });
  }
  if (!verifiedRows) {
    ui.alert("No verified rows for " + idea.id + " in " + rdName + ".\nRun ⑤ Publish first."); return;
  }

  // 2) Scan the Visual Library scenes for this idea.
  const vlSh = ss.getSheetByName(typeof SHEET !== "undefined" ? SHEET.VISUAL : "Visual Library");
  if (!vlSh) { ui.alert("No Visual Library tab."); return; }
  const vl = vlSh.getDataRange().getValues();

  const problems = [];
  let sceneCount = 0, scenesWithData = 0;
  for (let i = 1; i < vl.length; i++) {
    if (String(vl[i][0]).trim() !== idea.id) continue;
    sceneCount++;
    const sceneNum   = vl[i][1];
    const remotionData = vl[i][18] == null ? "" : String(vl[i][18]).trim();
    if (remotionData) scenesWithData++;

    const haystack = [ rbStripRenderParams_(rbInlineCounterUnit_(vl[i][18])), vl[i][15], vl[i][16], vl[i][3] ]   // REMOTION_DATA, CP_EVENT, CP_ANGLE, DESCRIPTION
      .map(function (x) { return x == null ? "" : String(x); }).join("  ");

    const unverified = rbNumberTokens_(haystack).filter(function (t) {
      return !rbTokenVerified_(t, verified) && !rbBenignNumber_(t);
    });
    const uniq = unverified.filter(function (t, k) { return unverified.indexOf(t) === k; });
    if (uniq.length) problems.push("Scene " + sceneNum + ": " + uniq.join(", "));
  }

  // Empty render data is a FAILURE, not a pass — the gate must not green a blank video.
  if (sceneCount === 0) { ui.alert("No scenes for " + idea.id + " in Visual Library. Run Stage 4."); return; }
  if (scenesWithData === 0) {
    ui.alert("🚫 NO RENDER DATA",
      sceneCount + " scenes exist for " + idea.id + " but REMOTION_DATA is EMPTY on all of them.\n\n" +
      "Stage 4B did not fill the render columns. Re-run Stage 4B — Director Review, " +
      "then run this check again.", ui.ButtonSet.OK);
    return;
  }

  if (!problems.length) {
    ui.alert("✅ SCENE NUMBERS VERIFIED",
      "Every figure on screen for " + idea.id + " traces to a verified claim.\n\n" +
      scenesWithData + " of " + sceneCount + " scenes carry render data; " +
      verifiedRows + " verified rows checked.", ui.ButtonSet.OK);
    return;
  }

  ui.alert("🚫 UNVERIFIED FIGURES ON SCREEN",
    "These numbers appear in scenes but NOT in the verified evidence for " + idea.id + ":\n\n" +
    problems.join("\n") +
    "\n\nEeither they are fabricated (fix the scene) or the claim wasn't published " +
    "(re-run ⑤). Do not render until this is clean.",
    ui.ButtonSet.OK);
}

// Canonical numeric tokens so "$100 million", "$100M", "100,000,000" all compare
// as one value. Magnitude words AND single-letter suffixes (M/B/K/T) normalise to
// the same unit. The single-letter form uses a negative lookahead so the "b" in
// "banks" is never mistaken for "billion".
function rbNumberTokens_(s) {
  if (s == null) return [];
  const UNIT = { million:"m", billion:"b", trillion:"t", thousand:"k",
                 percent:"pct", "%":"pct", m:"m", b:"b", t:"t", k:"k" };
  const MULT = { m:1e6, b:1e9, t:1e12, k:1e3 };
  const out = [];
  // Number = comma-grouped thousands ("1,534,280") OR a plain run ("2100000"),
  // then optional decimal. The (?!\d) stops "50,2016" from merging into one token.
  const re = /\$?\s?((?:\d{1,3}(?:,\d{3})+(?!\d)|\d+)(?:\.\d+)?)\s*(million|billion|trillion|thousand|percent|%|[mbkt](?![a-z]))?/gi;
  let m;
  while ((m = re.exec(String(s))) !== null) {
    const num  = m[1].replace(/,/g, "");
    const unit = UNIT[(m[2] || "").toLowerCase()] || "";
    out.push(num + unit);
    // Also emit the expanded integer so "2.1 million" and "2100000" (the counter's
    // raw value) compare equal — otherwise a verified figure trips the gate.
    if (MULT[unit]) out.push(String(Math.round(parseFloat(num) * MULT[unit])));
  }
  return out;
}

// COUNTER_ANIMATION splits a magnitude figure across two fields
// ("to=35 | unit=¥B+"), so the target number reaches the tokenizer BARE and can
// never match the verified "35b" built from "35 billion yen" — a false block on
// a properly sourced figure. Rewrite the pair so the number and its magnitude
// are adjacent ("to=35 B"), letting the normal tokenizer emit the right token.
// The unit must still match: a bare number with no verified counterpart at that
// magnitude is left bare and still blocks.
/* Some REMOTION_DATA keys are RENDER PARAMETERS, not claims: how long the scene
   is on screen, where a counter starts, which checkpoint index this is. They were
   being tokenised and demanded as evidence, so `duration=15` blocked a scene for
   an unverified "15" — a scene-length setting the director chose, nothing to do
   with the case. (`from=0` only slipped through because 0 is benign; any duration
   over 12 seconds would block.) Strip them before the evidence scan. */
function rbStripRenderParams_(s) {
  return String(s == null ? "" : s)
    .replace(/\b(duration|from|num|total|failure_index|splitfrac|delay)\s*=\s*[^|]*/gi, "");
}

// EVERY array format splits the magnitude away from the number the same way
// COUNTER_ANIMATION does:
//     to=35 | unit=¥B+
//     points=Zi-A Capital:60:true | unit=$M
//     gauges=Independent directors:11:%:one of nine:true
// so the tokenizer sees a bare "60" while only "60m" is verified, and the gate
// blocks a properly sourced figure. Re-attach the unit to each value so the
// normal tokenizer emits the right token. The unit must still MATCH — a bare
// number with no verified counterpart at that magnitude is left bare and blocks.
function rbInlineCounterUnit_(remotionData) {
  const s = String(remotionData == null ? "" : remotionData);
  const unit = s.match(/\bunit\s*=\s*([^|]+)/i);

  // Only a REAL magnitude counts: a word (million/billion/…), a percent sign, or
  // a lone M/B/K/T once currency decoration is stripped ("¥B+" → "B"). A label
  // like unit=accounts must NOT be read as trillion just because it contains "t".
  var suffix = "";
  if (unit) {
    const u = String(unit[1]).trim();
    const word = u.match(/\b(million|billion|trillion|thousand|percent)\b/i);
    const bare = u.replace(/[^a-z]/gi, "");
    suffix = word            ? " " + word[1]
           : /%/.test(u)     ? " percent"
           : /^[mbkt]$/i.test(bare) ? " " + bare
           : "";
  }

  var out = s;

  // COUNTER_ANIMATION — "to=35" → "to=35 B"
  if (suffix && /\bto\s*=/i.test(out)) {
    out = out.replace(/\bto\s*=\s*([\d.,]+)/i, function (w, num) { return "to=" + num + suffix; });
  }

  // OPENING_HOOK / DATA_CALLOUT / STAT_POSTER — "value=140 | unit=$M+" → "value=140 M".
  // Missed before, so an opening hook blocked on a bare "140" while "140m" was
  // verified. (The "+" in "$M+" is stripped by the magnitude test above.)
  if (suffix && /\bvalue\s*=/i.test(out)) {
    out = out.replace(/\bvalue\s*=\s*([$¥€£]?\s*[\d.,]+)/i, function (w, num) { return "value=" + num + suffix; });
  }

  // BAR_CHART / LINE_GRAPH — "points=Label:60:true" → "points=Label:60 M:true"
  if (suffix && /\bpoints\s*=/i.test(out)) {
    out = out.replace(/\bpoints\s*=\s*([^|]+)/i, function (w, list) {
      return "points=" + list.replace(/:\s*([\d.,]+)(?=\s*(?::|,|$))/g, function (w2, num) { return ":" + num + suffix; });
    });
  }

  // PROGRESS_GAUGE — the unit is the THIRD field of each item
  // ("Independent directors:11:%:one of nine:true"), so read it per item.
  if (/\bgauges\s*=/i.test(out)) {
    out = out.replace(/\bgauges\s*=\s*([^|]+)/i, function (w, list) {
      return "gauges=" + list.split(",").map(function (item) {
        const p = item.split(":");
        if (p.length >= 3 && /^[\d.,]+$/.test(String(p[1]).trim())) {
          const gu = String(p[2] || "").trim();
          const gs = /%/.test(gu) ? " percent"
                   : /^[mbkt]$/i.test(gu.replace(/[^a-z]/gi, "")) ? " " + gu.replace(/[^a-z]/gi, "")
                   : "";
          if (gs) p[1] = String(p[1]).trim() + gs;
        }
        return p.join(":");
      }).join(",");
    });
  }

  return out;
}

/* "$750K" and "$750,000" are the SAME figure, but they tokenise differently:
   the first yields BOTH "750k" and "750000", the second only "750000". The
   whitelist holds whichever form the source document used, so a scene written
   with a K/M/B suffix was blocked by evidence written in full — and vice versa.
   Accept either representation of the same value. This does NOT loosen the gate:
   a figure with no verified counterpart at any representation still blocks. */
function rbTokenVerified_(tok, verified) {
  if (verified[tok]) return true;
  const MULT = { k: 1e3, m: 1e6, b: 1e9, t: 1e12 };

  // "750k" → is "750000" verified?
  const m = String(tok).match(/^([\d.]+)([kmbt])$/);
  if (m && MULT[m[2]]) {
    if (verified[String(Math.round(parseFloat(m[1]) * MULT[m[2]]))]) return true;
  }

  // "750000" → is "750k" (or 0.75m …) verified?
  if (/^\d+$/.test(tok)) {
    const n = parseInt(tok, 10);
    for (const u in MULT) {
      if (n >= MULT[u]) {
        const q = n / MULT[u];
        // both "15m" and "0.1m" style forms
        if (verified[String(q) + u]) return true;
        if (verified[String(+q.toFixed(3)) + u]) return true;
      }
    }
  }
  return false;
}

// Structural numbers that aren't claims: years, small pure integers (ordinals /
// scene counts), and CFPA statute citations. A magnitude-suffixed token like
// "3b" is NEVER benign — that's the class we must catch.
function rbBenignNumber_(tok) {
  if (/^(19|20)\d{2}$/.test(tok)) return true;                       // a year
  if (/^\d+$/.test(tok) && parseInt(tok, 10) <= 12) return true;     // small ordinal / count, no unit
  if (/^(1031|1036|5481|5531|5536|5563|5565)$/.test(tok)) return true;  // CFPA statute sections
  return false;
}


// ── ④ Validate before render (the gate) ──────────────────────────────────────
function validateBeforeRender() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dmSh = ss.getSheetByName(SHEET_RB.DM);
  const clSh = ss.getSheetByName(SHEET_RB.CLAIMS);
  if (!dmSh || dmSh.getLastRow() < 2) { ui.alert("🚫 BLOCKED", "No data moments.", ui.ButtonSet.OK); return; }

  // Index the claims by id.
  const claims = {};
  clSh.getRange(2, 1, clSh.getLastRow() - 1, COL_RB_CLAIMS.length).getValues()
    .forEach(function (r) { claims[r[1]] = { status: r[2], tier: r[3], conflictWith: r[14] }; });

  const dm = dmSh.getRange(2, 1, dmSh.getLastRow() - 1, COL_RB_DM.length).getValues();
  const problems = [], conflicted = [];

  dm.forEach(function (r) {
    const dmId = r[0], claimId = r[1];
    const c = claims[claimId];
    if (!c)                       { problems.push(dmId + ": claim " + claimId + " not found"); return; }
    if (c.status !== "Verified")  { problems.push(dmId + ": " + claimId + " is " + c.status); }
    if (c.tier === "Boilerplate") { problems.push(dmId + ": " + claimId + " is boilerplate"); }
    if (c.conflictWith)           { conflicted.push(dmId + " (" + claimId + ") conflicts with " + c.conflictWith); }
  });

  if (problems.length) {
    ui.alert("🚫 RENDER BLOCKED", problems.join("\n") +
      "\n\nEvery on-screen number must map to a Verified, non-boilerplate claim.", ui.ButtonSet.OK);
    return;
  }

  ui.alert("✅ EVIDENCE GATE PASSED",
    dm.length + " data moment(s), all Verified and document-backed." +
    (conflicted.length
      ? "\n\n⚠ These figures have a conflicting counterpart — show the source and date on screen:\n" +
        conflicted.join("\n")
      : ""),
    ui.ButtonSet.OK);
}
