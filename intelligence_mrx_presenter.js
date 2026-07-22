/* ============================================================================
   Intelligence_3.2_MrXPresenter.gs — GovernX Intelligence Platform
   PHASE 3 · UNIT 3.2 — Stage 7C: Mr. X Presenter + Director Expression Logic

   Two parts:
   1. buildCharacterBlueprint(videoId) — the DIRECTOR pass. Reads a video's
      Visual Library scenes and assigns Mr. X's expression/gesture/position per
      scene, using SCENE_EXPRESSION_MAP. Pure logic, no API, no cost.
   2. generateMrXPresenter(videoId) — STAGE 7C. Takes the ElevenLabs voiceover +
      the chosen expression PNG and renders a lip-synced host clip via a
      talking-head vendor. VENDOR-AGNOSTIC: the actual API call sits behind
      MRX_LIPSYNC_VENDOR, so choosing D-ID or HeyGen later is a small wiring step.

   PIPELINE POSITION:  Script → Voiceover (7) → Mr. X Animation (7C) → Assembly (9B)

   SAFETY: additive, own constants, no config.gs edits. Reads production tabs
   (Visual Library, Publishing Tracker) via COL_VISUAL / COL_PUBLISHING from
   config.gs; writes the blueprint to the intelligence Character_Blueprint tab.

   HOW TO USE:
   1. Select a video row (Visual Library / Publishing Tracker / Idea Catalogue),
      run  buildCharacterBlueprint()  → fills Character_Blueprint for that video.
   2. Later, once art + a vendor are set: run  generateMrXPresenter().
   ============================================================================ */


// ── Lip-sync vendor config ───────────────────────────────────────────────────
// Both vendors are wired. Flip this one word to switch: "heygen" | "did".
const MRX_LIPSYNC_VENDOR     = "did";
const MRX_PRESENTER_POSITION = "corner-BR";
const MRX_WIDTH     = 1080;   // output dimensions (tune for your corner overlay)
const MRX_HEIGHT    = 1080;
const MRX_POLL_MS   = 15000;  // 15s between status checks
const MRX_MAX_POLLS = 20;     // ~5 min cap (Apps Script hard limit is 6 min)
//
// Script Properties required (Project Settings → Script Properties):
//   HEYGEN_API_KEY  +  HEYGEN_TALKING_PHOTO_ID   (for HeyGen)
//   DID_API_KEY                                   (for D-ID)
// IMPORTANT: the voiceover audio URL — and, for D-ID, the Mr. X image URL — must
// be PUBLICLY reachable. Share the Drive files "anyone with the link".


// ── Director rules: production scene type → Mr. X expression ─────────────────
// Data-driven so you can tune without touching logic.
const SCENE_EXPRESSION_MAP = {
  "opening title" : "Neutral",
  "title"         : "Neutral",
  "text"          : "Serious",
  "checkpoint"    : "Thinking",
  "timeline"      : "Thinking",
  "infographic"   : "Neutral",     // overridden to Warning if it's a counter/risk
  "risk matrix"   : "Warning",
  "kpi dashboard" : "Neutral",
  "gauge"         : "Warning",
  "data table"    : "Neutral"
};

// Per-expression default staging
const EXPRESSION_STAGING = {
  "Neutral"   : { gesture: "steady, hands settled", emotion: "composed" },
  "Serious"   : { gesture: "still, direct to camera", emotion: "grave" },
  "Thinking"  : { gesture: "hand near chin", emotion: "analytical" },
  "Warning"   : { gesture: "raised index finger", emotion: "alarmed" },
  "Surprised" : { gesture: "slight recoil, brows up", emotion: "astonished" },
  "Happy"     : { gesture: "open palm, slight nod", emotion: "assured" }
};

const INTEL_HEADER_BG_PR = "#1a1a2e";
const INTEL_HEADER_FG_PR = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// PART 1 — DIRECTOR PASS: build the per-scene Mr. X blueprint for a video
// ══════════════════════════════════════════════════════════════════════════════
function buildCharacterBlueprint(videoId) {
  const prodSS = SpreadsheetApp.getActiveSpreadsheet();  // production tabs
  const ss     = intelSS_();                              // intelligence tabs
  const ui     = SpreadsheetApp.getUi();

  if (!videoId) videoId = getSelectedVideoId_();
  if (!videoId) return;

  const visual = prodSS.getSheetByName(SHEET.VISUAL);    // SHEET/COL_VISUAL from config.gs
  if (!visual) { ui.alert("Visual Library tab not found."); return; }

  const bp = ss.getSheetByName(SHEET_MRX.BLUEPRINT);      // from Unit 3.1
  if (!bp) { ui.alert("Run setupMrXTabs() first (Character_Blueprint missing)."); return; }

  const vData = visual.getDataRange().getValues();
  const rows  = [];
  for (let i = 1; i < vData.length; i++) {
    if ((vData[i][COL_VISUAL.ID - 1] || "").toString().trim() !== videoId) continue;

    const sceneNum  = vData[i][COL_VISUAL.SCENE_NUM - 1];
    const sceneType = (vData[i][COL_VISUAL.SCENE_TYPE - 1] || "").toString().trim();
    const remotion  = (vData[i][COL_VISUAL_EXTENDED.REMOTION_DATA - 1] || "").toString().toLowerCase();

    const expression = mapExpressionForScene_(sceneType, remotion);
    const stage      = EXPRESSION_STAGING[expression] || EXPRESSION_STAGING["Neutral"];
    const sceneId    = videoId + "-S" + sceneNum;

    rows.push([sceneId, expression, stage.gesture, stage.emotion, MRX_PRESENTER_POSITION, "idle-talk"]);
  }

  if (!rows.length) { ui.alert("No scenes found in Visual Library for " + videoId + "."); return; }

  // Clean rebuild for this video
  const bpData = bp.getDataRange().getValues();
  for (let i = bpData.length - 1; i >= 1; i--) {
    if ((bpData[i][COL_MRX_BLUEPRINT.SCENE_ID - 1] || "").toString().indexOf(videoId + "-S") === 0) {
      bp.deleteRow(i + 1);
    }
  }
  bp.getRange(bp.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

  ui.alert("✅ Character Blueprint built for " + videoId,
    rows.length + " scenes directed.\n\n" +
    "Expression mix: " + summarizeExpressions_(rows), ui.ButtonSet.OK);
}

// Scene type (+ remotion data) → expression
function mapExpressionForScene_(sceneType, remotionLower) {
  const key = sceneType.toLowerCase();
  // Counter animations / risk cues override to Warning
  if (remotionLower.indexOf("counter") !== -1 || remotionLower.indexOf("risk") !== -1) return "Warning";
  return SCENE_EXPRESSION_MAP[key] || "Neutral";
}

function summarizeExpressions_(rows) {
  const counts = {};
  rows.forEach(r => { counts[r[1]] = (counts[r[1]] || 0) + 1; });
  return Object.keys(counts).map(k => k + ":" + counts[k]).join("  ");
}


// ══════════════════════════════════════════════════════════════════════════════
// PART 2 — STAGE 7C: render the lip-synced Mr. X presenter (vendor-agnostic)
// ══════════════════════════════════════════════════════════════════════════════
function generateMrXPresenter(videoId) {
  const prodSS = SpreadsheetApp.getActiveSpreadsheet();
  const ss     = intelSS_();
  const ui     = SpreadsheetApp.getUi();

  if (!videoId) videoId = getSelectedVideoId_();
  if (!videoId) return;

  // ── Preflight: vendor configured? ─────────────────────────────────────────
  if (!MRX_LIPSYNC_VENDOR) {
    ui.alert("Stage 7C — vendor not set",
      "Lip-sync is not wired to a provider yet.\n\n" +
      "1. Choose D-ID or HeyGen.\n" +
      "2. Set MRX_LIPSYNC_VENDOR = \"did\" or \"heygen\" in this file.\n" +
      "3. Add DID_API_KEY / HEYGEN_API_KEY to Script Properties.\n" +
      "4. Fill the MrX_Library PNG column with the 6 expression images.\n\n" +
      "The Director pass (buildCharacterBlueprint) works now without any of this.",
      ui.ButtonSet.OK);
    return;
  }

  // ── Gather inputs: voiceover audio + base expression PNG ──────────────────
  const audioUrl = readPublishingField_(prodSS, videoId, COL_PUBLISHING.VOICEOVER_AUDIO);
  if (!audioUrl) { ui.alert("No voiceover audio found for " + videoId + " (run Stage 7 first)."); return; }

  const baseExpression = dominantExpression_(ss, videoId) || "Neutral";
  const pngUrl = readMrXPng_(ss, baseExpression);
  if (!pngUrl) { ui.alert("No PNG for expression '" + baseExpression + "' in MrX_Library."); return; }

  ui.alert("Stage 7C — Generating Mr. X",
    "Vendor: " + MRX_LIPSYNC_VENDOR + "\nExpression: " + baseExpression +
    "\nThis calls the talking-head API and may take 1-3 minutes.", ui.ButtonSet.OK);

  try {
    const clipUrl = lipSyncRender_(audioUrl, pngUrl);   // vendor switch below

    // Persist the render — the vendor URL is temporary (expires ~24h)
    const saved = saveMrXClipToDrive_(videoId, baseExpression, clipUrl);
    logMrXRender_(videoId, baseExpression, saved.driveUrl);

    ui.alert("✅ Mr. X presenter rendered & saved",
      "Saved to Drive (permanent):\n" + saved.driveUrl +
      "\n\nLogged in the MrX_Renders tab. Use it as a corner overlay in Stage 9B assembly.",
      ui.ButtonSet.OK);
  } catch (err) {
    if (typeof logError === "function") logError("Stage 7C — Mr. X", videoId, "API/Runtime", err.message);
    ui.alert("❌ Stage 7C failed: " + err.message);
  }
}

// ── Helper: list your HeyGen talking photos + their IDs (run once) ───────────
// Uploads happen in the HeyGen app; this prints each talking_photo_id so you can
// copy the right one into the HEYGEN_TALKING_PHOTO_ID Script Property.
function heyGenListTalkingPhotos() {
  const ui  = SpreadsheetApp.getUi();
  const key = PropertiesService.getScriptProperties().getProperty("HEYGEN_API_KEY");
  if (!key) { ui.alert("Add HEYGEN_API_KEY to Script Properties first."); return; }

  const resp = UrlFetchApp.fetch("https://api.heygen.com/v1/talking_photo.list", {
    method: "get", headers: { "x-api-key": key, "accept": "application/json" },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) { ui.alert("HeyGen error " + resp.getResponseCode() + ":\n" + resp.getContentText()); return; }

  const json = JSON.parse(resp.getContentText() || "{}");
  // Find the array of talking photos regardless of exact wrapper shape
  let list = [];
  if (Array.isArray(json.data)) list = json.data;
  else if (json.data && Array.isArray(json.data.talking_photos)) list = json.data.talking_photos;
  else if (json.data && Array.isArray(json.data.list)) list = json.data.list;

  if (!list.length) { ui.alert("No talking photos found. Upload Mr. X in the HeyGen app first, then rerun."); return; }

  const lines = list.map(tp => {
    const id   = tp.talking_photo_id || tp.id || "?";
    const name = tp.talking_photo_name || tp.name || "(unnamed)";
    return name + "  →  " + id;
  });
  Logger.log("HeyGen talking photos:\n" + lines.join("\n"));
  ui.alert("Your HeyGen Talking Photos",
    lines.join("\n") + "\n\nCopy the right ID into Script Properties → HEYGEN_TALKING_PHOTO_ID.",
    ui.ButtonSet.OK);
}

// ── Persist a rendered clip to Drive + log it ─────────────────────────────────
const MRX_CLIPS_FOLDER = "GovernX Mr. X Clips";

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

// Download the vendor's temporary clip and save it permanently to Drive.
// Returns { driveUrl, fileId }.
function saveMrXClipToDrive_(videoId, expression, sourceUrl) {
  const resp = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true });
  if (resp.getResponseCode() >= 300)
    throw new Error("Could not download rendered clip (" + resp.getResponseCode() + ").");

  const name = videoId + "_mrx_" + String(expression || "neutral").toLowerCase() + ".mp4";
  const blob = resp.getBlob().setName(name);
  const file = getOrCreateFolder_(MRX_CLIPS_FOLDER).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* ok */ }
  return { driveUrl: file.getUrl(), fileId: file.getId() };
}

// Append a row to the MrX_Renders log (creates the tab on first use).
function logMrXRender_(videoId, expression, driveUrl) {
  const ss = intelSS_();
  let t = ss.getSheetByName("MrX_Renders");
  if (!t) {
    t = ss.insertSheet("MrX_Renders");
    t.getRange(1, 1, 1, 4).setValues([["Video_ID", "Expression", "Clip_URL", "Rendered_At"]])
     .setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
    t.setFrozenRows(1);
    [200, 130, 380, 160].forEach((w, i) => t.setColumnWidth(i + 1, w));
    if (typeof INTEL_TAB_COLOR !== "undefined") t.setTabColor(INTEL_TAB_COLOR);
  }
  // Upsert: one latest render per video
  const data = t.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][0] || "").toString().trim() === videoId) t.deleteRow(i + 1);
  }
  t.appendRow([videoId, expression, driveUrl, new Date()]);
}

// ── Vendor switch — the ONE place to implement each provider ──────────────────
function lipSyncRender_(audioUrl, imageUrl) {
  switch (MRX_LIPSYNC_VENDOR) {
    case "did":    return lipSyncDID_(audioUrl, imageUrl);
    case "heygen": return lipSyncHeyGen_(audioUrl, imageUrl);
    default: throw new Error("Unknown MRX_LIPSYNC_VENDOR: " + MRX_LIPSYNC_VENDOR);
  }
}

// D-ID: create a talk from an image + audio, then poll until done.
// Docs: POST https://api.d-id.com/talks  ·  GET https://api.d-id.com/talks/{id}
function lipSyncDID_(audioUrl, imageUrl) {
  const key = PropertiesService.getScriptProperties().getProperty("DID_API_KEY");
  if (!key) throw new Error("DID_API_KEY missing from Script Properties.");
  if (!imageUrl) throw new Error("D-ID needs a public Mr. X image URL (from MrX_Library PNG).");

  // D-ID keys come as "email:code" — base64 that. If a bare code was given, append ":".
  const auth = "Basic " + Utilities.base64Encode(key.indexOf(":") !== -1 ? key : key + ":");

  // D-ID rejects Google Drive "view" links (source_url must end in .png/.jpg and
  // audio must be a direct media URL). So upload both to D-ID first and use the
  // hosted URLs it returns. Works with normal Drive share links.
  const didImageUrl = uploadAssetToDID_(driveBlobFromUrl_(imageUrl), "images", "image", auth);
  const didAudioUrl = uploadAssetToDID_(driveBlobFromUrl_(audioUrl), "audios", "audio", auth);

  const create = UrlFetchApp.fetch("https://api.d-id.com/talks", {
    method: "post", contentType: "application/json",
    headers: { "Authorization": auth, "accept": "application/json" },
    payload: JSON.stringify({
      source_url: didImageUrl,
      script: { type: "audio", audio_url: didAudioUrl },
      config: { stitch: true }
    }),
    muteHttpExceptions: true
  });
  if (create.getResponseCode() >= 300)
    throw new Error("D-ID create failed " + create.getResponseCode() + ": " + create.getContentText());
  const id = (JSON.parse(create.getContentText() || "{}")).id;
  if (!id) throw new Error("D-ID: no talk id returned.");

  for (let i = 0; i < MRX_MAX_POLLS; i++) {
    Utilities.sleep(MRX_POLL_MS);
    const st = UrlFetchApp.fetch("https://api.d-id.com/talks/" + id, {
      method: "get", headers: { "Authorization": auth, "accept": "application/json" },
      muteHttpExceptions: true
    });
    const sj = JSON.parse(st.getContentText() || "{}");
    if (sj.status === "done" && sj.result_url) return sj.result_url;
    if (sj.status === "error" || sj.status === "rejected")
      throw new Error("D-ID render error: " + (sj.error ? JSON.stringify(sj.error) : st.getContentText()));
  }
  throw new Error("D-ID timed out after " + MRX_MAX_POLLS + " polls (Apps Script 6-min limit).");
}

// Extract a Google Drive file id from any Drive URL form
function driveIdFromUrl_(url) {
  const s = String(url || "");
  let m = s.match(/\/d\/([-\w]+)/);        if (m) return m[1];   // .../file/d/<id>/view
  m     = s.match(/[?&]id=([-\w]+)/);      if (m) return m[1];   // ...?id=<id>
  m     = s.match(/[-\w]{25,}/);           return m ? m[0] : ""; // fallback: long token
}

// Get the file's bytes as a Blob from a Drive URL (needs Drive scope — already granted)
function driveBlobFromUrl_(url) {
  const id = driveIdFromUrl_(url);
  if (!id) throw new Error("Could not read a Drive file id from URL: " + url);
  return DriveApp.getFileById(id).getBlob();
}

// Upload a blob to D-ID's asset store; returns the hosted URL.
// kind = "images" | "audios",  field = "image" | "audio"
// D-ID validates the file name/extension, so we force a valid one + content type.
function uploadAssetToDID_(blob, kind, field, auth) {
  const isImg = (kind === "images");
  const ct = (blob.getContentType() || "").toLowerCase();

  let ext, mime;
  if (isImg) {
    if (ct.indexOf("jpeg") !== -1 || ct.indexOf("jpg") !== -1) { ext = ".jpg"; mime = "image/jpeg"; }
    else                                                       { ext = ".png"; mime = "image/png"; }
  } else {
    if      (ct.indexOf("wav") !== -1) { ext = ".wav"; mime = "audio/wav"; }
    else if (ct.indexOf("mp4") !== -1) { ext = ".mp4"; mime = "audio/mp4"; }
    else if (ct.indexOf("m4a") !== -1 || ct.indexOf("aac") !== -1) { ext = ".m4a"; mime = "audio/mp4"; }
    else                               { ext = ".mp3"; mime = "audio/mpeg"; }
  }

  const named = blob.copyBlob();
  named.setName((isImg ? "mrx_image" : "mrx_audio") + ext);
  named.setContentType(mime);

  const payload = {}; payload[field] = named;
  const resp = UrlFetchApp.fetch("https://api.d-id.com/" + kind, {
    method: "post",
    headers: { "Authorization": auth, "accept": "application/json" },
    payload: payload,                     // object with a named Blob → multipart/form-data
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300)
    throw new Error("D-ID " + kind + " upload failed " + resp.getResponseCode() + ": " + resp.getContentText());
  const j = JSON.parse(resp.getContentText() || "{}");
  if (!j.url) throw new Error("D-ID " + kind + " upload returned no url: " + resp.getContentText());
  return j.url;
}

// HeyGen: generate from a pre-uploaded talking photo + audio, then poll.
// Docs: POST https://api.heygen.com/v2/video/generate
//       GET  https://api.heygen.com/v1/video_status.get?video_id=...
// Upload the Mr. X photo once in HeyGen, then store its id in HEYGEN_TALKING_PHOTO_ID.
function lipSyncHeyGen_(audioUrl, imageUrl) {
  const props = PropertiesService.getScriptProperties();
  const key   = props.getProperty("HEYGEN_API_KEY");
  const tpId  = props.getProperty("HEYGEN_TALKING_PHOTO_ID");
  if (!key)  throw new Error("HEYGEN_API_KEY missing from Script Properties.");
  if (!tpId) throw new Error("HEYGEN_TALKING_PHOTO_ID missing — upload Mr. X once in HeyGen and paste its talking_photo_id into Script Properties.");

  const create = UrlFetchApp.fetch("https://api.heygen.com/v2/video/generate", {
    method: "post", contentType: "application/json",
    headers: { "x-api-key": key, "accept": "application/json" },
    payload: JSON.stringify({
      video_inputs: [{
        character: { type: "talking_photo", talking_photo_id: tpId },
        voice    : { type: "audio", audio_url: audioUrl }
      }],
      dimension: { width: MRX_WIDTH, height: MRX_HEIGHT }
    }),
    muteHttpExceptions: true
  });
  const cj = JSON.parse(create.getContentText() || "{}");
  if (create.getResponseCode() >= 300 || cj.error)
    throw new Error("HeyGen create failed: " + create.getContentText());
  const vid = cj.data && cj.data.video_id;
  if (!vid) throw new Error("HeyGen: no video_id returned.");

  for (let i = 0; i < MRX_MAX_POLLS; i++) {
    Utilities.sleep(MRX_POLL_MS);
    const st = UrlFetchApp.fetch("https://api.heygen.com/v1/video_status.get?video_id=" + vid, {
      method: "get", headers: { "x-api-key": key, "accept": "application/json" },
      muteHttpExceptions: true
    });
    const sj = JSON.parse(st.getContentText() || "{}");
    const status = sj.data && sj.data.status;
    if (status === "completed") return sj.data.video_url;
    if (status === "failed")
      throw new Error("HeyGen render failed: " + JSON.stringify(sj.data && sj.data.error));
  }
  throw new Error("HeyGen timed out after " + MRX_MAX_POLLS + " polls (Apps Script 6-min limit).");
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
// Read the ID from the selected row (works on any tab whose ID is column A)
function getSelectedVideoId_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui    = SpreadsheetApp.getUi();
  const row   = sheet.getActiveCell().getRow();
  if (row < 2) { ui.alert("Select a data row of a video (ID in column A)."); return null; }
  const id = (sheet.getRange(row, 1).getValue() || "").toString().trim();
  if (!/^GX-/.test(id)) { ui.alert("Selected row doesn't have a GX- video ID in column A."); return null; }
  return id;
}

function dominantExpression_(ss, videoId) {
  const bp = ss.getSheetByName(SHEET_MRX.BLUEPRINT);
  if (!bp) return null;
  const data = bp.getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_MRX_BLUEPRINT.SCENE_ID - 1] || "").toString().indexOf(videoId + "-S") === 0) {
      const e = data[i][COL_MRX_BLUEPRINT.EXPRESSION - 1];
      if (e) counts[e] = (counts[e] || 0) + 1;
    }
  }
  let best = null, max = 0;
  Object.keys(counts).forEach(k => { if (counts[k] > max) { max = counts[k]; best = k; } });
  return best;
}

function readMrXPng_(ss, expression) {
  const lib = ss.getSheetByName(SHEET_MRX.LIBRARY);
  if (!lib) return "";
  const data = lib.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_MRX_LIB.EXPRESSION - 1] || "").toString().trim() === expression) {
      return (data[i][COL_MRX_LIB.PNG - 1] || "").toString().trim();
    }
  }
  return "";
}

function readPublishingField_(prodSS, videoId, col) {
  const pub = prodSS.getSheetByName(SHEET.PUBLISHING);
  if (!pub) return "";
  const data = pub.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_PUBLISHING.ID - 1] || "").toString().trim() === videoId) {
      const raw = (data[i][col - 1] || "").toString().trim();
      if (/^https?:/.test(raw)) return raw;
      const formula = pub.getRange(i + 1, col).getFormula();
      const m = formula.match(/HYPERLINK\("(https?:[^"]+)"/i);
      return m ? m[1] : raw;
    }
  }
  return "";
}
