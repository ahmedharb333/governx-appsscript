/* ============================================================================
   Stage_9C_Remotion_Assembly.gs — GovernX Content OS
   Stage 9C — Remotion Film Assembly (replaces Shotstack / Stage 9B)

   Sends the Visual Library scenes to the Remotion server's /assemble/job
   endpoint, which downloads each scene's public Stage-7B MP3, locks each
   scene to its real audio length, adapts REMOTION_DATA → case-file props,
   and renders ONE synced MP4 with the investigative case-file look.

   Because the render takes minutes (well past the 6-min Apps Script limit)
   this is a TWO-STEP flow:
     1. assembleFilmRemotion()      — submit the job, poll for a few minutes
     2. checkAssembleFilmRemotion() — re-runnable; poll again until done

   Mirrors Stage 9B for save behaviour: downloads the MP4, saves
   {id}_final_video.mp4 to getOrCreateContentFolder, writes the Drive link
   to the Publishing Tracker, resolves the Error Log.
   ============================================================================ */

const ASSEMBLE_JOB_PROP  = "REMOTION_ASSEMBLE_JOB";  // Script Property: JSON { contentId, jobId, submittedAt }
const ASSEMBLE_POLL_MS   = 10000;                    // poll interval
const ASSEMBLE_POLL_BUDGET_MS = 270000;              // ~4.5 min per run — stays under the 6-min GAS limit


// ══════════════════════════════════════════════════════════════════════════════
// STEP 1 — SUBMIT
// ══════════════════════════════════════════════════════════════════════════════
function assembleFilmRemotion() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Read scenes from Visual Library ────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  if (!visualSheet) { ui.alert("No Visual Library tab found."); return; }
  const visualData = visualSheet.getDataRange().getValues();

  const scenes       = [];
  const audioFileIds = [];

  for (let i = 1; i < visualData.length; i++) {
    const row = visualData[i];
    if (row[COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;

    const sceneNum     = row[COL_VISUAL.SCENE_NUM  - 1].toString().trim();
    const type         = row[COL_VISUAL.SCENE_TYPE - 1].toString().trim();
    const remotionData = row.length >= COL_VISUAL_EXTENDED.REMOTION_DATA
      ? row[COL_VISUAL_EXTENDED.REMOTION_DATA - 1].toString().trim() : "";
    const voiceSync    = row.length >= COL_VISUAL_EXTENDED.VOICEOVER_SYNC
      ? row[COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1].toString().trim() : "";
    let   audioUrl     = row.length >= COL_VISUAL_EXTENDED.VOICEOVER_AUDIO_URL
      ? row[COL_VISUAL_EXTENDED.VOICEOVER_AUDIO_URL - 1].toString().trim() : "";

    // A scene needs REMOTION_DATA to be renderable at all.
    if (!remotionData) continue;
    if (!audioUrl.startsWith("http")) audioUrl = "";

    const fileId = audioUrl ? extractDriveFileId(audioUrl) : null;
    if (fileId) audioFileIds.push(fileId);

    scenes.push({
      sceneNum      : sceneNum,
      type          : type,
      remotionData  : remotionData,
      voiceoverSync : voiceSync,
      audioUrl      : audioUrl
    });
  }

  if (scenes.length === 0) {
    ui.alert("No renderable scenes for: " + idea.id +
      "\n\nEach scene needs REMOTION_DATA (col 19). Run Stage 4 / 4B first.");
    return;
  }

  const withAudio = scenes.filter(s => s.audioUrl).length;

  const confirm = ui.alert("🎬 Stage 9C — Remotion Film Assembly",
    "Assemble one synced film for: " + idea.company + "\n\n" +
    "Scenes           : " + scenes.length + "\n" +
    "Per-scene audio  : " + withAudio + "/" + scenes.length +
      (withAudio === scenes.length ? "  (✅ full frame-accurate sync)"
       : withAudio === 0 ? "  (⚠️ none — run Stage 7B for synced timing)"
       : "  (⚠️ some scenes will use a default length — run Stage 7B)") + "\n" +
    "Captions         : off (film style)\n\n" +
    "Step 1: Make scene audio public + submit render job\n" +
    "Step 2: This run polls ~4.5 min. If the render isn't done,\n" +
    "        use “Check Assembly Status” to poll again.\n\n" +
    "Proceed?", ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  try {
    // ── Make each scene's audio publicly downloadable (server pulls via uc?export=download)
    let madePublic = 0;
    audioFileIds.forEach(function(fileId) {
      try {
        DriveApp.getFileById(fileId)
          .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        madePublic++;
      } catch (e) {
        Logger.log("Could not make audio public (non-fatal): " + fileId + " — " + e.message);
      }
    });
    Logger.log("Stage 9C: made " + madePublic + "/" + audioFileIds.length + " audio files public");

    // ── Submit the job ─────────────────────────────────────────────────────────
    const serverBase = getRemotionServerUrl();
    const res = UrlFetchApp.fetch(serverBase + "/assemble/job", {
      method            : "post",
      contentType       : "application/json",
      headers           : { "ngrok-skip-browser-warning": "true" },
      payload           : JSON.stringify({
        contentId    : idea.id,
        scenes       : scenes,
        showCaptions : true          // bold narration subtitle band on every scene
      }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = JSON.parse(res.getContentText());
    if (code !== 200 || !body.ok || !body.jobId) {
      throw new Error("Submit failed (" + code + "): " + (body.error || res.getContentText()).toString().substring(0, 300));
    }

    const jobId = body.jobId;
    PropertiesService.getScriptProperties().setProperty(ASSEMBLE_JOB_PROP,
      JSON.stringify({ contentId: idea.id, jobId: jobId, submittedAt: Date.now() }));
    updatePipelineStatus_(idea.id, "S9", "⏳");
    Logger.log("Stage 9C: submitted job " + jobId + " for " + idea.id);

    // ── Poll within this run's budget ──────────────────────────────────────────
    const result = pollAssembleJob_(serverBase, jobId, ASSEMBLE_POLL_BUDGET_MS);

    if (result.status === "done") {
      finalizeAssembly_(idea, serverBase, result);
    } else if (result.status === "error") {
      updatePipelineStatus_(idea.id, "S9", "❌");
      logError("Stage 9C — Remotion Assembly", idea.id, "Render Error", result.error);
      ui.alert("❌ Stage 9C Failed", result.error + "\nSee Error Log.", ui.ButtonSet.OK);
    } else {
      ui.alert("⏳ Still rendering",
        "Job submitted (" + jobId + ") and still rendering after ~4.5 min.\n\n" +
        "Leave the Remotion server running, wait a few minutes, then run:\n" +
        "GovernX → 🎬 Check Assembly Status\n\n" +
        "You can run Check as many times as needed.",
        ui.ButtonSet.OK);
    }

  } catch (err) {
    updatePipelineStatus_(idea.id, "S9", "❌");
    logError("Stage 9C — Remotion Assembly", idea.id, "Submit Error", err.message);
    ui.alert("❌ Stage 9C Failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// STEP 2 — CHECK (re-runnable)
// ══════════════════════════════════════════════════════════════════════════════
function checkAssembleFilmRemotion() {

  const ui  = SpreadsheetApp.getUi();
  const raw = PropertiesService.getScriptProperties().getProperty(ASSEMBLE_JOB_PROP);
  if (!raw) {
    ui.alert("No assembly job in progress.\nRun 🎬 Assemble Film (Remotion) first.");
    return;
  }

  let job;
  try { job = JSON.parse(raw); } catch (e) { job = null; }
  if (!job || !job.jobId || !job.contentId) {
    ui.alert("Stored job is unreadable. Re-run 🎬 Assemble Film (Remotion).");
    PropertiesService.getScriptProperties().deleteProperty(ASSEMBLE_JOB_PROP);
    return;
  }

  // Rebuild a minimal idea for saving (contentId + company).
  const idea = ideaFromContentId_(job.contentId);

  try {
    const serverBase = getRemotionServerUrl();
    const result = pollAssembleJob_(serverBase, job.jobId, ASSEMBLE_POLL_BUDGET_MS);

    if (result.status === "done") {
      finalizeAssembly_(idea, serverBase, result);
    } else if (result.status === "error") {
      updatePipelineStatus_(idea.id, "S9", "❌");
      logError("Stage 9C — Remotion Assembly", idea.id, "Render Error", result.error);
      ui.alert("❌ Assembly Failed", result.error + "\nSee Error Log.", ui.ButtonSet.OK);
      PropertiesService.getScriptProperties().deleteProperty(ASSEMBLE_JOB_PROP);
    } else {
      const mins = Math.round((Date.now() - (job.submittedAt || Date.now())) / 60000);
      ui.alert("⏳ Still rendering",
        "Job " + job.jobId + " is still rendering (~" + mins + " min elapsed).\n\n" +
        "Wait a bit longer, then run 🎬 Check Assembly Status again.",
        ui.ButtonSet.OK);
    }

  } catch (err) {
    logError("Stage 9C — Remotion Assembly", idea.id, "Check Error", err.message);
    ui.alert("❌ Check Failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// POLL — returns { status: "done"|"running"|"error", ... } within maxMs
// ══════════════════════════════════════════════════════════════════════════════
function pollAssembleJob_(serverBase, jobId, maxMs) {
  const deadline = Date.now() + maxMs;
  let last = { status: "running" };

  while (Date.now() < deadline) {
    Utilities.sleep(ASSEMBLE_POLL_MS);
    const res = UrlFetchApp.fetch(serverBase + "/assemble/job/" + jobId, {
      method            : "get",
      headers           : { "ngrok-skip-browser-warning": "true" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) continue;

    let body;
    try { body = JSON.parse(res.getContentText()); } catch (e) { continue; }

    if (body.status === "done") return Object.assign({ status: "done" }, body);
    if (body.status === "error" || body.ok === false) {
      return { status: "error", error: body.error || "unknown render error" };
    }
    last = { status: "running", elapsedMs: body.elapsedMs };
    Logger.log("Stage 9C: job " + jobId + " running… " +
      (body.elapsedMs ? Math.round(body.elapsedMs / 1000) + "s" : ""));
  }
  return last;
}


// ══════════════════════════════════════════════════════════════════════════════
// FINALIZE — download MP4, save to Drive, link in Publishing Tracker
// ══════════════════════════════════════════════════════════════════════════════
function finalizeAssembly_(idea, serverBase, result) {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const filename = idea.id + "_final_video.mp4";

  // PREFER the file the SERVER already streamed straight to Drive — no download
  // here, so no memory blowup at any file size. Apps Script only pulls the file
  // itself as a FALLBACK (Drive not configured on the server), and that path
  // holds the whole video as a number[] (~8x its size) and OOMs on large films.
  // To keep the server-upload path working: service-account.json + DRIVE_FOLDER_ID
  // on the render server (see governx-remotion/src/server/drive-upload.js).
  let driveUrl;
  if (result.drive && result.drive.driveUrl) {
    driveUrl = result.drive.driveUrl;
    Logger.log("Stage 9C: server uploaded to Drive — " + driveUrl);
  } else {
    if (result.driveError) {
      Logger.log("Stage 9C: server Drive upload failed (" + result.driveError + ") — falling back to Apps Script download.");
    }
    // result.url = http://localhost:3000/output/{filename} — rewrite to the public server base.
    const mp4Url = String(result.url || "")
      .replace(/https?:\/\/localhost:\d+/, serverBase)
      .replace(/https?:\/\/127\.0\.0\.1:\d+/, serverBase);
    Logger.log("Stage 9C: downloading MP4 from " + mp4Url);
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
    const blob          = downloadRenderedFile_(mp4Url, filename, result.bytes);
    driveUrl            = contentFolder.createFile(blob).getUrl();
  }

  // Write the final-video link to the Publishing Tracker (mirror Stage 9B).
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (pubSheet) {
    const pubData = pubSheet.getDataRange().getValues();
    for (let i = 1; i < pubData.length; i++) {
      if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
        pubSheet.getRange(i + 1, COL_PUBLISHING.SCENES_FOLDER).setValue(driveUrl);
        break;
      }
    }
  }

  // Write REAL per-scene start times back to the Visual Library TIMESTAMP column.
  // Stage 10 builds YouTube chapters from this column, which previously held
  // meaningless time-of-day serials; now it holds each scene's actual position in
  // the assembled film (M:SS), so chapters finally line up with the video.
  try { writeSceneTimestamps_(idea.id, result.sceneTimings); } catch (e) { Logger.log("Stage 9C: timestamp write skipped — " + e.message); }

  updatePipelineStatus_(idea.id, "S9", "✅");
  SpreadsheetApp.flush();
  try { autoResolveErrorLog(idea.id); } catch (e) {}
  PropertiesService.getScriptProperties().deleteProperty(ASSEMBLE_JOB_PROP);

  ui.alert("✅ Stage 9C Complete",
    "Film assembled and saved:\n\n" + driveUrl + "\n\n" +
    "Scenes   : " + (result.sceneCount || "?") + "\n" +
    "Duration : " + (result.durationSec != null ? result.durationSec + "s" : "?") + "\n\n" +
    "Ready for YouTube. Run Stage 10 for metadata.",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPER — write per-scene start times (from the assembly) to Visual Library
// TIMESTAMP column, matched by SCENE_NUM. Real film positions for Stage 10.
// ══════════════════════════════════════════════════════════════════════════════
function writeSceneTimestamps_(contentId, sceneTimings) {
  if (!sceneTimings || !sceneTimings.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.VISUAL);
  if (!sh || sh.getLastRow() < 2) return;

  // sceneNum → "M:SS"
  const byNum = {};
  sceneTimings.forEach(function (t) { byNum[String(t.sceneNum).trim()] = t.mmss; });

  const data = sh.getDataRange().getValues();
  let wrote = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][COL_VISUAL.ID - 1]).trim() !== contentId) continue;
    const mmss = byNum[String(data[i][COL_VISUAL.SCENE_NUM - 1]).trim()];
    if (mmss === undefined) continue;
    // Force TEXT so Sheets does not coerce "1:23" into a duration/time serial.
    sh.getRange(i + 1, COL_VISUAL.TIMESTAMP).setNumberFormat("@").setValue(mmss);
    wrote++;
  }
  Logger.log("Stage 9C: wrote " + wrote + " scene timestamps");
}

// ══════════════════════════════════════════════════════════════════════════════
// HELPER — rebuild a minimal idea {id, company, language} from a content ID
// so Check can save without an active Idea Catalogue selection.
// ══════════════════════════════════════════════════════════════════════════════
function ideaFromContentId_(contentId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let company = contentId, language = "English";
  try {
    const ideaSheet = ss.getSheetByName(SHEET.IDEA);
    if (ideaSheet) {
      const data = ideaSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][COL_IDEA.ID - 1].toString().trim() === contentId) {
          company  = data[i][COL_IDEA.COMPANY - 1] || contentId;
          language = data[i][COL_IDEA.LANGUAGE_FLAG - 1] || "English";
          break;
        }
      }
    }
  } catch (e) { Logger.log("ideaFromContentId_: " + e.message); }
  return { id: contentId, company: company, language: language };
}


/* ── Downloading a render that is bigger than one UrlFetchApp response ────────
   UrlFetchApp caps a response at 50 MB and TRUNCATES silently past it — no error,
   no warning, just a short file. A 56 MB film landed in Drive 419 bytes under
   exactly 50 MiB and played 36 seconds short; the render itself was byte-perfect,
   the loss happened entirely in transit and nothing noticed.

   So: ask the server how big the file is, pull it in HTTP Range chunks, and
   refuse to save anything whose byte count does not match. Better to fail loudly
   than to publish a film that is quietly missing its ending.
   ---------------------------------------------------------------------------- */
const ASSEMBLE_CHUNK_BYTES = 20 * 1024 * 1024;   // well inside the 50 MB cap

function downloadRenderedFile_(mp4Url, filename, expectedBytes) {
  const headers = { "ngrok-skip-browser-warning": "true" };

  // Prefer the size the render reported; otherwise ask the server.
  let total = Number(expectedBytes) || 0;
  if (!total) {
    try {
      const infoUrl = mp4Url.replace("/output/", "/output-info/");
      const info = JSON.parse(UrlFetchApp.fetch(infoUrl, { headers: headers, muteHttpExceptions: true }).getContentText());
      if (info && info.ok) total = Number(info.bytes) || 0;
    } catch (e) { Logger.log("Stage 9C: size probe failed — " + e.message); }
  }

  // Small enough for one response: fetch whole, but still verify the length.
  if (total && total <= ASSEMBLE_CHUNK_BYTES) {
    const r = UrlFetchApp.fetch(mp4Url, { headers: headers, muteHttpExceptions: true });
    assertDownloadOk_(r, total);
    return r.getBlob().setContentType("video/mp4").setName(filename);
  }

  if (!total) {
    // Unknown size — a single fetch is the only option, so verify what we can.
    const r = UrlFetchApp.fetch(mp4Url, { headers: headers, muteHttpExceptions: true });
    assertDownloadOk_(r, 0);
    const bytes = r.getBlob().getBytes().length;
    if (bytes >= 50 * 1024 * 1024 - 4096) {
      throw new Error("Download hit the 50 MB UrlFetchApp cap and is almost certainly truncated (" +
        bytes + " bytes). The server render is fine — retry once the server reports its size.");
    }
    return r.getBlob().setContentType("video/mp4").setName(filename);
  }

  // Chunked path.
  let bytes = [];
  let start = 0;
  while (start < total) {
    const end = Math.min(start + ASSEMBLE_CHUNK_BYTES - 1, total - 1);
    const res = UrlFetchApp.fetch(mp4Url, {
      muteHttpExceptions: true,
      headers: { "ngrok-skip-browser-warning": "true", "Range": "bytes=" + start + "-" + end }
    });
    const code = res.getResponseCode();
    if (code !== 206 && code !== 200) {
      throw new Error("Range request failed (HTTP " + code + ") for bytes " + start + "-" + end +
        ".\nThe server must support Range — restart it so the updated /output route is live.");
    }
    const part = res.getBlob().getBytes();
    if (code === 200) {                      // server ignored Range and sent everything
      if (part.length !== total) {
        throw new Error("Server ignored Range and the single response is " + part.length +
          " of " + total + " bytes — truncated.");
      }
      bytes = part;
      break;
    }
    bytes = bytes.concat(part);
    Logger.log("Stage 9C: " + bytes.length + " / " + total + " bytes");
    start = end + 1;
  }

  if (bytes.length !== total) {
    throw new Error("INCOMPLETE DOWNLOAD: got " + bytes.length + " of " + total +
      " bytes. Nothing was saved — the render on the server is intact, so just retry.");
  }
  Logger.log("Stage 9C: downloaded " + bytes.length + " bytes (verified)");
  return Utilities.newBlob(bytes, "video/mp4", filename);
}

function assertDownloadOk_(res, expectedBytes) {
  const code = res.getResponseCode();
  const ct = (res.getHeaders()["content-type"] || "").toLowerCase();
  if (code !== 200 && code !== 206) {
    throw new Error("Could not download the assembled MP4 (HTTP " + code + ").\n" +
      "Make sure the Remotion server + ngrok are running and REMOTION_SERVER_URL is current.");
  }
  if (ct.indexOf("text") !== -1 || ct.indexOf("html") !== -1) {
    throw new Error("Server returned " + ct + " instead of video — check the server and ngrok.");
  }
  if (expectedBytes) {
    const got = res.getBlob().getBytes().length;
    if (got !== expectedBytes) {
      throw new Error("INCOMPLETE DOWNLOAD: got " + got + " of " + expectedBytes + " bytes.");
    }
  }
}

/* ── Preview One Scene ─────────────────────────────────────────────────────────
   Renders a SINGLE Visual Library scene the way the ASSEMBLY does: its
   REMOTION_DATA is routed through the server's adapter (/preview-scene → adapt.js)
   into the real case-file component, then rendered to a still. The old preview
   used the legacy /render compositions (CheckpointCard / InfographicScene / …),
   which is why it looked nothing like the finished film. */
function previewOneScene() {
  const ui   = SpreadsheetApp.getUi();
  const idea = getActiveIdeaRow();
  if (!idea) return;
  if (!checkRendererHealth(ui)) return;

  const vs   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.VISUAL);
  const data = vs.getDataRange().getValues();

  const scenes = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const rd = data[i].length >= COL_VISUAL_EXTENDED.REMOTION_DATA
      ? data[i][COL_VISUAL_EXTENDED.REMOTION_DATA - 1].toString().trim() : "";
    if (rd.length < 5) continue;
    scenes.push({
      num : data[i][COL_VISUAL.SCENE_NUM  - 1].toString().trim(),
      type: data[i][COL_VISUAL.SCENE_TYPE - 1].toString().trim(),
      rd  : rd
    });
  }
  if (!scenes.length) {
    ui.alert("No scenes with REMOTION_DATA for " + idea.id + ".\nRun Stage 4 / 4B first.");
    return;
  }

  const list = scenes.map(function (s) { return "Scene " + s.num + "  (" + s.type + ")"; }).join("\n");
  const resp = ui.prompt("🔍 Preview One Scene — " + idea.company,
    "Scenes with data:\n\n" + list + "\n\nEnter a scene number to preview (renders ~1 min):",
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const target = scenes.find(function (s) { return s.num === resp.getResponseText().trim(); });
  if (!target) { ui.alert("Scene not found. Check the number and try again."); return; }

  try {
    const serverBase = getRemotionServerUrl();
    const res = UrlFetchApp.fetch(serverBase + "/preview-scene", {
      method: "post", contentType: "application/json",
      headers: { "ngrok-skip-browser-warning": "true" },
      payload: JSON.stringify({ sceneType: target.type, remotionData: target.rd }),
      muteHttpExceptions: true
    });
    let body;
    try { body = JSON.parse(res.getContentText()); }
    catch (e) { throw new Error("Server did not return JSON (HTTP " + res.getResponseCode() +
      "). Is the render engine running? Use 🩺 Check render engine is reachable."); }
    if (res.getResponseCode() !== 200 || !body.success) {
      throw new Error(body.error || "Preview failed (HTTP " + res.getResponseCode() + ")");
    }

    const pngUrl = String(body.url)
      .replace(/https?:\/\/localhost:\d+/, serverBase)
      .replace(/https?:\/\/127\.0\.0\.1:\d+/, serverBase);
    const png = UrlFetchApp.fetch(pngUrl, { muteHttpExceptions: true, headers: { "ngrok-skip-browser-warning": "true" } });
    const bytes = png.getResponseCode() === 200 ? png.getBlob().getBytes().length : 0;
    if (bytes < 200) throw new Error("Could not download the preview (HTTP " + png.getResponseCode() + ", " + bytes + " bytes).");

    const folder = getOrCreateContentFolder(idea.id, idea.company);
    const name   = idea.id + "_preview_scene_" + target.num + ".png";
    const existing = folder.getFilesByName(name);
    while (existing.hasNext()) existing.next().setTrashed(true);
    const file = folder.createFile(png.getBlob().setName(name).setContentType("image/png"));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    ui.alert("✅ Scene " + target.num + " preview",
      "Component: " + body.component + "  ← the real case-file component the film uses\n\n" +
      file.getUrl() + "\n\nSaved to the content folder. Open the link to view.",
      ui.ButtonSet.OK);
  } catch (err) {
    logError("Preview One Scene", idea.id, "Preview Error", err.message);
    ui.alert("❌ Preview failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}
