/* ============================================================================
   Stage_9B_Shotstack.gs — GovernX Content OS
   Stage 9B — Shotstack Automated Video Assembly
   ============================================================================ */

const SHOTSTACK_EDIT_URL   = "https://api.shotstack.io/edit/v1/render";
const SHOTSTACK_POLL_URL   = "https://api.shotstack.io/edit/v1/render/";
const SHOTSTACK_INGEST_URL = "https://api.shotstack.io/ingest/v1/sources";
const SHOTSTACK_INGEST_POLL= "https://api.shotstack.io/ingest/v1/sources/";
const SHOTSTACK_FPS        = 25;
const POLL_INTERVAL_MS     = 8000;
const MAX_POLLS            = 60;

const GX_BLACK = "#0A0A0A";
const GX_RED   = "#FF0000";
const GX_WHITE = "#FFFFFF";


// ══════════════════════════════════════════════════════════════════════════════
// MAIN — Stage 9B
// ══════════════════════════════════════════════════════════════════════════════
function renderWithShotstack() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui     = SpreadsheetApp.getUi();
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey = PropertiesService.getScriptProperties()
                   .getProperty("SHOTSTACK_API_KEY");

  if (!apiKey) {
    ui.alert("⚠️ Shotstack API Key Missing",
      "Add your key to Script Properties:\n" +
      "Key: SHOTSTACK_API_KEY\nValue: your_production_key_from_shotstack.io",
      ui.ButtonSet.OK);
    return;
  }

  // ── Read scenes from Visual Library ────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  for (let i = 1; i < visualData.length; i++) {
    const row = visualData[i];
    if (row[COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;

    const aiClipUrl = row[COL_VISUAL.AI_CLIP_URL - 1].toString().trim();
    const link      = row[COL_VISUAL.LINK        - 1].toString().trim();
    const clipUrl   = aiClipUrl || link;

    // Skip empty or non-URL values
    if (!clipUrl) continue;
    if (!clipUrl.startsWith("http://") && !clipUrl.startsWith("https://")) {
      Logger.log("Skipping placeholder scene " +
        row[COL_VISUAL.SCENE_NUM - 1] + ": " + clipUrl.substring(0, 60));
      continue;
    }

    // Bug 4 guard: warn if clip belongs to a different content ID
    const clipFilename   = clipUrl.split("/").pop().split("?")[0];
    const foreignIdMatch = clipFilename.match(/GX-\d{4}-[A-Z]+-\d{3}/);
    if (foreignIdMatch && foreignIdMatch[0] !== idea.id) {
      Logger.log("⚠️  Bug 4 warning — Scene " +
        row[COL_VISUAL.SCENE_NUM - 1] + ": clip filename contains foreign ID " +
        foreignIdMatch[0] + " (expected " + idea.id + "). URL: " + clipUrl);
    }

    // Read VOICEOVER_SYNC (col 21) and per-scene audio URL (col 23)
    const voiceSync     = (typeof COL_VISUAL_EXTENDED !== "undefined" && row.length >= 21)
      ? row[COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1].toString().trim()
      : (row.length >= 21 ? row[20].toString().trim() : "");
    const sceneAudioUrl = (typeof COL_VISUAL_EXTENDED !== "undefined" && row.length >= 23)
      ? row[COL_VISUAL_EXTENDED.VOICEOVER_AUDIO_URL - 1].toString().trim()
      : "";

    scenes.push({
      num           : row[COL_VISUAL.SCENE_NUM   - 1].toString().trim(),
      type          : row[COL_VISUAL.SCENE_TYPE  - 1].toString().trim(),
      desc          : row[COL_VISUAL.DESCRIPTION - 1].toString().trim(),
      timestamp     : row[COL_VISUAL.TIMESTAMP   - 1].toString().trim(),
      clipUrl       : clipUrl,
      voiceSync     : voiceSync,
      sceneAudioUrl : sceneAudioUrl.startsWith("http") ? sceneAudioUrl : ""
    });
  }

  if (scenes.length === 0) {
    ui.alert("No clips found for: " + idea.id +
      "\nRun Stages 8A, 8B, 8C, 8D first.");
    return;
  }

  // ── Read voiceover URL ──────────────────────────────────────────────────────
  let audioUrl = "";
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (pubSheet) {
    const pubData = pubSheet.getDataRange().getValues();
    for (let i = 1; i < pubData.length; i++) {
      if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
        const rawAudio = pubData[i][COL_PUBLISHING.VOICEOVER_AUDIO - 1].toString().trim();
        if (rawAudio.startsWith("http")) {
          audioUrl = rawAudio;
        } else {
          const formula = pubSheet.getRange(i + 1, COL_PUBLISHING.VOICEOVER_AUDIO).getFormula();
          const urlMatch = formula.match(/HYPERLINK\("(https?:[^"]+)"/i);
          if (urlMatch) {
            audioUrl = urlMatch[1];
            Logger.log("Audio URL from formula: " + audioUrl);
          }
        }
        break;
      }
    }
  }

  const master = getMasterContent(idea.id);

  // ── Confirm ─────────────────────────────────────────────────────────────────
  const syncCount       = scenes.filter(function(s) { return s.voiceSync && s.voiceSync.trim() !== ""; }).length;
  const perSceneAudio   = scenes.filter(function(s) { return s.sceneAudioUrl !== ""; }).length;
  const usePerSceneMode = perSceneAudio >= Math.ceil(scenes.length * 0.5); // ≥50% of scenes have audio

  const confirm = ui.alert("🎬 Stage 9B — Shotstack Render",
    "Ready to assemble: " + idea.company + "\n\n" +
    "Clips            : " + scenes.length + "\n" +
    "Per-scene audio  : " + perSceneAudio + "/" + scenes.length + " scenes (" +
      (usePerSceneMode ? "✅ Using per-scene mode — precise sync" : "⚠️ Fewer than 50% — falling back to global audio") + ")\n" +
    "Global voiceover : " + (audioUrl ? "✅ Found (used as fallback)" : "⚠️ Not found") + "\n\n" +
    (usePerSceneMode
      ? "TIMING MODE: Per-scene audio\nEach scene's duration = its actual MP3 duration from Stage 7B.\nThis gives frame-accurate audio/video sync.\n\n"
      : "TIMING MODE: Character-count proportional\nRun Stage 7B first to unlock per-scene precision.\n\n") +
    "Step 1: Ingest clips to Shotstack (~1 min)\n" +
    "Step 2: Render final video (~3 min)\n" +
    "Step 3: Save MP4 to Drive\n\n" +
    "Proceed?", ui.ButtonSet.YES_NO);

  if (confirm !== ui.Button.YES) return;

  try {
    ui.alert("⏳ Step 1 of 3 — Preparing clips",
      "Making " + scenes.length + " clips publicly accessible...\n" +
      "Then uploading to Shotstack (~1–2 min).\nClick OK to start.",
      ui.ButtonSet.OK);

    makeClipsPublic(scenes, audioUrl);
    const ingestedScenes = ingestAllClips(scenes, audioUrl, apiKey, usePerSceneMode);

    ui.alert("⏳ Step 2 of 3 — Rendering",
      "All clips ingested. Submitting render job...\n" +
      "Rendering takes 2–4 minutes.\nClick OK to start.",
      ui.ButtonSet.OK);

    ingestedScenes.scenes.forEach(s => {
      Logger.log("Scene " + s.num + " URL: " + s.clipUrl +
        (s.ingestedAudioUrl ? " | Audio: " + s.ingestedAudioUrl + " (" + (s.audioDurationSec || "?") + "s)" : ""));
    });
    Logger.log("Global audio URL for render: " + ingestedScenes.audioUrl);

    const timeline = buildTimeline(ingestedScenes.scenes,
                                   ingestedScenes.audioUrl,
                                   ingestedScenes.audioDuration,
                                   master, idea, usePerSceneMode);

    const renderId = submitRender(timeline, apiKey);
    Logger.log("Stage 9B: Render ID: " + renderId);

    const renderUrl = pollRender(renderId, apiKey);
    if (!renderUrl) throw new Error("Render timed out. Check Shotstack dashboard.");

    const filename      = idea.id + "_final_video.mp4";
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);

    const videoRes = UrlFetchApp.fetch(renderUrl, { muteHttpExceptions: true });
    if (videoRes.getResponseCode() !== 200) {
      throw new Error("Could not download rendered video from Shotstack.");
    }

    const driveFile = contentFolder.createFile(videoRes.getBlob().setName(filename));
    const driveUrl  = driveFile.getUrl();

    if (pubSheet) {
      const pubData = pubSheet.getDataRange().getValues();
      for (let i = 1; i < pubData.length; i++) {
        if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
          pubSheet.getRange(i + 1, COL_PUBLISHING.SCENES_FOLDER).setValue(driveUrl);
          break;
        }
      }
    }

    SpreadsheetApp.flush();
    try { autoResolveErrorLog(idea.id); } catch(e) {}

    ui.alert("✅ Stage 9B Complete",
      "Final video assembled and saved:\n\n" +
      driveUrl + "\n\n" +
      "Ready for YouTube upload.\nRun Stage 10 for metadata.",
      ui.ButtonSet.OK);

  } catch (err) {
    logError("Stage 9B — Shotstack", idea.id, "Render Error", err.message);
    ui.alert("❌ Stage 9B Failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// INGEST ALL CLIPS
// When usePerSceneMode=true, also ingests per-scene audio (col 23) for each
// scene that has a VOICEOVER_AUDIO_URL. Returns per-scene ingestedAudioUrl and
// audioDurationSec so buildTimeline can use exact MP3 durations for timing.
// ══════════════════════════════════════════════════════════════════════════════
function ingestAllClips(scenes, audioUrl, apiKey, usePerSceneMode) {

  const ingestJobs = [];

  scenes.forEach((scene, i) => {
    const publicUrl = convertDriveUrl(scene.clipUrl);

    if (!publicUrl || !publicUrl.startsWith("http")) {
      Logger.log("Skipping invalid URL for scene " + scene.num + ": " + publicUrl);
      ingestJobs.push({ sceneIndex: i, sourceId: null, type: "video", fallbackUrl: null });
      return;
    }

    // Drive files: binary blob upload directly to Shotstack (bypasses redirect)
    if (scene.clipUrl.includes("drive.google.com") || extractDriveFileId(scene.clipUrl)) {
      try {
        const fileId = extractDriveFileId(scene.clipUrl);
        const file   = DriveApp.getFileById(fileId);
        const blob   = file.getBlob();
        const bytes  = blob.getBytes();
        Logger.log("Uploading Drive blob scene " + scene.num + ": " + file.getName() + " (" + bytes.length + " bytes)");

        const initRes  = UrlFetchApp.fetch("https://api.shotstack.io/ingest/v1/upload", {
          method           : "post",
          contentType      : "application/json",
          headers          : { "x-api-key": apiKey, "Accept": "application/json" },
          payload          : JSON.stringify({ filename: file.getName() }),
          muteHttpExceptions: true
        });
        const initCode = initRes.getResponseCode();
        const initBody = JSON.parse(initRes.getContentText());
        Logger.log("Upload init scene " + scene.num + " (" + initCode + "): " + JSON.stringify(initBody).substring(0, 200));

        if ((initCode === 200 || initCode === 201) && initBody.data && initBody.data.attributes && initBody.data.attributes.url) {
          const putRes = UrlFetchApp.fetch(initBody.data.attributes.url, {
            method           : "put",
            contentType      : blob.getContentType() || "video/mp4",
            payload          : bytes,
            muteHttpExceptions: true,
            followRedirects  : true
          });
          Logger.log("Upload PUT scene " + scene.num + ": " + putRes.getResponseCode());
          const sourceId = (initBody.data && initBody.data.id) ? initBody.data.id
                         : (initBody.data && initBody.data.attributes && initBody.data.attributes.id) ? initBody.data.attributes.id
                         : null;
          if (sourceId) {
            ingestJobs.push({ sceneIndex: i, sourceId: sourceId, type: "video" });
            Logger.log("Scene " + scene.num + " blob upload sourceId: " + sourceId);
            return;
          }
        }
      } catch (e) {
        Logger.log("Blob upload failed scene " + scene.num + " — " + e.message + " — falling back to usercontent URL");
      }

      const usercontent = "https://drive.usercontent.google.com/download?id=" +
        extractDriveFileId(scene.clipUrl) + "&export=download&authuser=0";
      ingestJobs.push({ sceneIndex: i, sourceId: null, type: "video", fallbackUrl: usercontent });
      return;
    }

    // Non-Drive URLs (Pexels, KlingAI, S3, etc)
    const response = UrlFetchApp.fetch(SHOTSTACK_INGEST_URL, {
      method           : "post",
      contentType      : "application/json",
      headers          : { "x-api-key": apiKey, "Accept": "application/json" },
      payload          : JSON.stringify({ url: publicUrl }),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const body = JSON.parse(response.getContentText());

    if (code === 201 || code === 200) {
      const sourceId = (body.data && body.data.id) ? body.data.id : body.id;
      ingestJobs.push({ sceneIndex: i, sourceId: sourceId, type: "video" });
      Logger.log("Ingest submitted scene " + scene.num + ": " + sourceId);
    } else {
      Logger.log("Ingest failed for scene " + scene.num + " (" + code + ") — using direct URL");
      ingestJobs.push({ sceneIndex: i, sourceId: null, type: "video", fallbackUrl: publicUrl });
    }
    Utilities.sleep(300);
  });

  // ── Per-scene audio ingest (Stage 7B MP3s from col 23) ─────────────────────
  const sceneAudioJobs = []; // { sceneIndex, sourceId }
  if (usePerSceneMode) {
    scenes.forEach((scene, i) => {
      if (!scene.sceneAudioUrl) return;
      try {
        const audioFileId = extractDriveFileId(scene.sceneAudioUrl);
        if (!audioFileId) return;
        const audioFile  = DriveApp.getFileById(audioFileId);
        const audioBlob  = audioFile.getBlob();
        const audioBytes = audioBlob.getBytes();
        Logger.log("Uploading scene " + scene.num + " audio: " + audioFile.getName() +
          " (" + audioBytes.length + " bytes)");

        const initRes  = UrlFetchApp.fetch("https://api.shotstack.io/ingest/v1/upload", {
          method           : "post",
          contentType      : "application/json",
          headers          : { "x-api-key": apiKey, "Accept": "application/json" },
          payload          : JSON.stringify({ filename: audioFile.getName() }),
          muteHttpExceptions: true
        });
        const initBody = JSON.parse(initRes.getContentText());
        if ((initRes.getResponseCode() === 200 || initRes.getResponseCode() === 201) &&
            initBody.data && initBody.data.attributes && initBody.data.attributes.url) {
          UrlFetchApp.fetch(initBody.data.attributes.url, {
            method           : "put",
            contentType      : audioBlob.getContentType() || "audio/mpeg",
            payload          : audioBytes,
            muteHttpExceptions: true,
            followRedirects  : true
          });
          const sourceId = initBody.data.id ||
            (initBody.data.attributes && initBody.data.attributes.id) || null;
          if (sourceId) {
            sceneAudioJobs.push({ sceneIndex: i, sourceId });
            Logger.log("Scene " + scene.num + " audio sourceId: " + sourceId);
          }
        }
      } catch (e) {
        Logger.log("Scene " + scene.num + " audio upload failed (non-fatal): " + e.message);
      }
      Utilities.sleep(200);
    });
  }

  // Audio ingest
  let audioSourceId = null;
  if (audioUrl && audioUrl !== "") {
    try {
      const audioFileId = extractDriveFileId(audioUrl);
      if (audioFileId) {
        const audioFile  = DriveApp.getFileById(audioFileId);
        const audioBlob  = audioFile.getBlob();
        const audioBytes = audioBlob.getBytes();
        Logger.log("Uploading audio blob: " + audioFile.getName() + " (" + audioBytes.length + " bytes)");

        const initRes  = UrlFetchApp.fetch("https://api.shotstack.io/ingest/v1/upload", {
          method           : "post",
          contentType      : "application/json",
          headers          : { "x-api-key": apiKey, "Accept": "application/json" },
          payload          : JSON.stringify({ filename: audioFile.getName() }),
          muteHttpExceptions: true
        });
        const initCode = initRes.getResponseCode();
        const initBody = JSON.parse(initRes.getContentText());

        if ((initCode === 200 || initCode === 201) && initBody.data && initBody.data.attributes && initBody.data.attributes.url) {
          const putRes = UrlFetchApp.fetch(initBody.data.attributes.url, {
            method          : "put",
            contentType     : audioBlob.getContentType() || "audio/mpeg",
            payload         : audioBytes,
            muteHttpExceptions: true,
            followRedirects : true
          });
          Logger.log("Audio PUT: " + putRes.getResponseCode());
          audioSourceId = (initBody.data && initBody.data.id) ? initBody.data.id
                        : (initBody.data && initBody.data.attributes && initBody.data.attributes.id) ? initBody.data.attributes.id
                        : null;
          Logger.log("Audio blob upload sourceId: " + audioSourceId);
        }
      } else {
        const audioRes = UrlFetchApp.fetch(SHOTSTACK_INGEST_URL, {
          method           : "post",
          contentType      : "application/json",
          headers          : { "x-api-key": apiKey, "Accept": "application/json" },
          payload          : JSON.stringify({ url: audioUrl }),
          muteHttpExceptions: true
        });
        if (audioRes.getResponseCode() === 201 || audioRes.getResponseCode() === 200) {
          const body = JSON.parse(audioRes.getContentText());
          audioSourceId = (body.data && body.data.id) ? body.data.id : body.id;
          Logger.log("Audio ingest submitted: " + audioSourceId);
        }
      }
    } catch (e) {
      Logger.log("Audio upload error: " + e.message);
    }
  }

  // Poll all jobs (video + global audio + per-scene audio)
  Logger.log("Polling ingest jobs...");
  const readyUrls      = {};
  const ingestDuration = {};

  // Collect all sourceIds that need polling
  const allSourceIds = [
    ...ingestJobs.filter(j => j.sourceId).map(j => j.sourceId),
    ...sceneAudioJobs.filter(j => j.sourceId).map(j => j.sourceId),
    ...(audioSourceId ? [audioSourceId] : [])
  ];

  for (let poll = 0; poll < 30; poll++) {
    Utilities.sleep(5000);
    let allReady = true;

    allSourceIds.forEach(function(sourceId) {
      if (readyUrls[sourceId]) return;
      const statusRes = UrlFetchApp.fetch(SHOTSTACK_INGEST_POLL + sourceId, {
        method : "get",
        headers: { "x-api-key": apiKey, "Accept": "application/json" },
        muteHttpExceptions: true
      });
      if (statusRes.getResponseCode() !== 200) { allReady = false; return; }

      const statusBody = JSON.parse(statusRes.getContentText());
      const status     = statusBody.data && statusBody.data.attributes ? statusBody.data.attributes.status : null;
      const url        = statusBody.data && statusBody.data.attributes
        ? (statusBody.data.attributes.url || statusBody.data.attributes.source) : null;

      Logger.log("Ingest " + sourceId + " status: " + status);

      if ((status === "ready" || status === "done") && url) {
        readyUrls[sourceId] = url;
        const dur = statusBody.data.attributes.duration;
        if (dur) ingestDuration[sourceId] = parseFloat(dur);
        Logger.log("Ingest ready: " + url + " duration: " + (dur || "unknown"));
      } else {
        allReady = false;
      }
    });

    if (allReady) break;
  }

  // Resolve video clips
  const updatedScenes = scenes.map(function(scene, i) {
    const videoJob    = ingestJobs.find(function(j) { return j.sceneIndex === i; });
    const ingestedUrl = videoJob && videoJob.sourceId ? readyUrls[videoJob.sourceId] : null;
    const shotUrl     = (ingestedUrl && ingestedUrl.startsWith("https://"))
      ? ingestedUrl : convertDriveUrl(scene.clipUrl);

    // Resolve per-scene audio
    const audioJob         = sceneAudioJobs.find(function(j) { return j.sceneIndex === i; });
    const ingestedAudioUrl = audioJob && audioJob.sourceId ? readyUrls[audioJob.sourceId] : null;
    const audioDurationSec = audioJob && audioJob.sourceId ? (ingestDuration[audioJob.sourceId] || null) : null;

    Logger.log("Scene " + scene.num + " final video: " + shotUrl +
      (ingestedAudioUrl ? " | audio: " + ingestedAudioUrl + " (" + (audioDurationSec || "?") + "s)" : ""));

    return Object.assign({}, scene, {
      clipUrl          : shotUrl,
      ingestedAudioUrl : ingestedAudioUrl || "",
      audioDurationSec : audioDurationSec
    });
  });

  const finalAudioUrl      = audioSourceId && readyUrls[audioSourceId]
    ? readyUrls[audioSourceId] : audioUrl;
  const finalAudioDuration = audioSourceId && ingestDuration[audioSourceId]
    ? ingestDuration[audioSourceId] : null;

  Logger.log("Global audio duration: " + (finalAudioDuration || "unknown — using defaults"));
  return { scenes: updatedScenes, audioUrl: finalAudioUrl, audioDuration: finalAudioDuration };
}


// ══════════════════════════════════════════════════════════════════════════════
// BUILD TIMELINE
// ══════════════════════════════════════════════════════════════════════════════
function buildTimeline(scenes, audioUrl, audioDuration, master, idea, usePerSceneMode) {

  const isArabic = idea.language === "Arabic" || idea.language === "Bilingual";

  // ── Scene timing ──────────────────────────────────────────────────────────
  // Per-scene mode: use actual MP3 duration (audioDurationSec) for each scene.
  // Fallback mode: proportional character-count timing (original behaviour).
  const sceneTimings = computeSceneTimings(scenes, audioDuration, usePerSceneMode);

  Logger.log("Scene timings: " + JSON.stringify(
    sceneTimings.map(function(s) {
      return {
        num  : s.num,
        start: s.startSec.toFixed(2),
        dur  : s.durationSec.toFixed(2),
        mode : s.audioDurationSec ? "exact-audio" : "proportional"
      };
    })
  ));

  // ── Track: video clips ────────────────────────────────────────────────────
  const videoClips = sceneTimings
    .filter(function(scene) {
      return scene.clipUrl &&
        typeof scene.clipUrl === "string" &&
        (scene.clipUrl.startsWith("http://") || scene.clipUrl.startsWith("https://"));
    })
    .map(function(scene) {
      return {
        asset     : { type: "video", src: scene.clipUrl, volume: 0 },
        start     : scene.startSec,
        length    : scene.durationSec,
        fit       : "cover",
        transition: { in: "fade" }
      };
    });

  if (videoClips.length === 0) {
    throw new Error("No valid clip URLs found. Check that your Drive files are shared publicly.");
  }

  // ── Track: per-scene audio (one clip per scene, timed to match video) ─────
  const audioClips = [];
  if (usePerSceneMode) {
    sceneTimings.forEach(function(scene) {
      if (!scene.ingestedAudioUrl) return;
      audioClips.push({
        asset : { type: "audio", src: scene.ingestedAudioUrl, volume: 1 },
        start : scene.startSec,
        length: scene.durationSec
      });
    });
    Logger.log("Per-scene audio clips: " + audioClips.length + "/" + sceneTimings.length);
  }

  // ── Track: Checkpoint overlays ────────────────────────────────────────────
  const textClips = sceneTimings
    .filter(function(s) { return s.type === "Checkpoint"; })
    .map(function(scene) {
      return {
        asset: {
          type    : "html",
          html    : buildCheckpointHtml(scene.desc, isArabic),
          width   : 1920,
          height  : 300,
          position: "bottom"
        },
        start     : scene.startSec,
        length    : scene.durationSec,
        position  : "bottom",
        offset    : { x: 0, y: 0.20 },
        transition: { in: "fade" }
      };
    });

  Logger.log("Tracks — videoClips: " + videoClips.length +
    " | audioClips: " + audioClips.length +
    " | checkpointOverlays: " + textClips.length);

  // Layer order: checkpoint overlays → audio → video
  const tracks = [];
  if (textClips.length > 0) tracks.push({ clips: textClips });
  if (audioClips.length > 0) tracks.push({ clips: audioClips });
  tracks.push({ clips: videoClips });

  const timeline = { background: GX_BLACK, tracks: tracks };

  // Global soundtrack: used in fallback mode, or as background music in per-scene mode
  // In per-scene mode, only add if no per-scene audio covers the full duration
  const uncoveredScenes = usePerSceneMode
    ? sceneTimings.filter(s => !s.ingestedAudioUrl && s.type !== "Timeline").length
    : sceneTimings.length;

  if (audioUrl && audioUrl !== "" && (!usePerSceneMode || uncoveredScenes > 0)) {
    timeline.soundtrack = {
      src   : audioUrl,
      effect: "fadeInFadeOut",
      volume: usePerSceneMode ? 0.15 : 1  // background-level in per-scene mode
    };
  }

  return {
    timeline: timeline,
    output  : { format: "mp4", resolution: "1080", fps: SHOTSTACK_FPS, quality: "high" }
  };
}


// ══════════════════════════════════════════════════════════════════════════════
// COMPUTE SCENE TIMINGS — replaces Gemini voice boundary detection
//
// APPROACH:
//   Each scene's duration is proportional to the character length of its
//   VOICEOVER_SYNC text relative to the total voiceover character count.
//   This mirrors real speech timing — more text = more time on screen.
//
// FIXED DURATION scenes (no voiceover content):
//   Timeline   → FIXED_TIMELINE_SEC  (closing scene, no VO)
//   Checkpoint → FIXED_CHECKPOINT_SEC (structural beat, brief VO window)
//
// FALLBACK: if a scene has no VOICEOVER_SYNC but is not a fixed type,
//   it receives an equal share of the remaining proportional budget.
//
// TOTAL DURATION: uses audioDuration from Shotstack ingest when available,
//   otherwise falls back to 90s (overridable in Script Properties).
// ══════════════════════════════════════════════════════════════════════════════
function computeSceneTimings(scenes, audioDuration, usePerSceneMode) {

  const FIXED_TIMELINE_SEC    = 8;   // Timeline scene — always 8s
  const FIXED_CHECKPOINT_SEC  = 4;   // Checkpoint scene — 4s punch-in
  const MIN_SCENE_SEC         = 2;   // floor: no scene shorter than 2s
  const MAX_SCENE_SEC         = 20;  // ceiling: cap runaway scenes
  const DISSOLVE_SEC          = 0.4; // overlap between consecutive scenes
  const AUDIO_TAIL_SEC        = 0.3; // extra buffer after audio ends before next scene

  // ── Step 1: establish total audio duration (fallback mode only) ───────────
  const totalAudioSec = (audioDuration && audioDuration > 5)
    ? audioDuration
    : 90;

  // ── Step 2: budget for proportional scenes (fallback mode) ───────────────
  var fixedBudget = 0;
  scenes.forEach(function(s) {
    if (s.type === "Timeline")   fixedBudget += FIXED_TIMELINE_SEC;
    if (s.type === "Checkpoint") fixedBudget += FIXED_CHECKPOINT_SEC;
  });
  const voiceoverBudget = Math.max(10, totalAudioSec - fixedBudget);

  // ── Step 3: character counts for proportional fallback ───────────────────
  var totalChars = 0;
  var noSyncCount = 0;
  scenes.forEach(function(s) {
    if (s.type === "Timeline" || s.type === "Checkpoint") return;
    const len = (s.voiceSync || "").length;
    if (len > 0) { totalChars += len; }
    else { noSyncCount++; }
  });
  const avgChars      = totalChars > 0 ? totalChars / Math.max(1, scenes.length - noSyncCount) : 80;
  const fallbackChars = avgChars * 0.6;
  const adjustedTotal = totalChars + (noSyncCount * fallbackChars);

  // ── Step 4: compute per-scene durations ──────────────────────────────────
  var runningStart = 0;
  const timings = scenes.map(function(scene) {
    var durationSec;

    if (scene.type === "Timeline") {
      // Fixed duration regardless of mode
      durationSec = FIXED_TIMELINE_SEC;

    } else if (scene.type === "Checkpoint") {
      durationSec = FIXED_CHECKPOINT_SEC;

    } else if (usePerSceneMode && scene.audioDurationSec && scene.audioDurationSec > 0) {
      // EXACT TIMING: use actual MP3 duration from ElevenLabs ingest + small tail
      durationSec = scene.audioDurationSec + AUDIO_TAIL_SEC;

    } else {
      // PROPORTIONAL FALLBACK: character-count share of voiceover budget
      const chars = (scene.voiceSync || "").length || fallbackChars;
      const share = adjustedTotal > 0 ? chars / adjustedTotal : 1 / scenes.length;
      durationSec = share * voiceoverBudget;
    }

    // Apply floor and ceiling
    durationSec = Math.min(MAX_SCENE_SEC, Math.max(MIN_SCENE_SEC, durationSec));

    // Dissolve overlap
    const startSec = runningStart > 0
      ? Math.max(0, runningStart - DISSOLVE_SEC)
      : 0;

    runningStart += durationSec;

    return Object.assign({}, scene, {
      startSec   : parseFloat(startSec.toFixed(2)),
      durationSec: parseFloat(durationSec.toFixed(2))
    });
  });

  const exactCount = usePerSceneMode
    ? scenes.filter(s => s.audioDurationSec && s.audioDurationSec > 0).length : 0;

  Logger.log("computeSceneTimings: mode=" + (usePerSceneMode ? "per-scene" : "proportional") +
    " | exact=" + exactCount + "/" + scenes.length +
    " | totalAudio=" + totalAudioSec.toFixed(1) + "s" +
    " | voiceoverBudget=" + voiceoverBudget.toFixed(1) + "s" +
    " | totalChars=" + Math.round(totalChars));

  return timings;
}


// ══════════════════════════════════════════════════════════════════════════════
// SUBMIT RENDER
// ══════════════════════════════════════════════════════════════════════════════
function submitRender(payload, apiKey) {
  const response = UrlFetchApp.fetch(SHOTSTACK_EDIT_URL, {
    method           : "post",
    contentType      : "application/json",
    headers          : { "x-api-key": apiKey, "Accept": "application/json" },
    payload          : JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const body = JSON.parse(response.getContentText());
  if (code !== 201 && code !== 200) {
    throw new Error("Shotstack submission failed (" + code + "): " + JSON.stringify(body));
  }
  const id = (body.response && body.response.id) ? body.response.id
           : (body.data && body.data.id) ? body.data.id : null;
  if (!id) throw new Error("No render ID returned: " + JSON.stringify(body));
  return id;
}


// ══════════════════════════════════════════════════════════════════════════════
// POLL RENDER STATUS
// ══════════════════════════════════════════════════════════════════════════════
function pollRender(renderId, apiKey) {
  for (var i = 0; i < MAX_POLLS; i++) {
    Utilities.sleep(POLL_INTERVAL_MS);
    const res = UrlFetchApp.fetch(SHOTSTACK_POLL_URL + renderId, {
      method : "get",
      headers: { "x-api-key": apiKey, "Accept": "application/json" },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) continue;
    const body   = JSON.parse(res.getContentText());
    const status = body.response ? body.response.status : null;
    const url    = body.response ? body.response.url    : null;
    Logger.log("Poll " + (i + 1) + ": " + status);
    if (status === "done" && url) return url;
    if (status === "failed") {
      const errDetail = body.response
        ? (body.response.error || body.response.message || body.response)
        : "unknown";
      throw new Error("Render failed: " + JSON.stringify(errDetail));
    }
  }
  return null;
}


// ══════════════════════════════════════════════════════════════════════════════
// MAKE CLIPS PUBLIC
// ══════════════════════════════════════════════════════════════════════════════
function makeClipsPublic(scenes, audioUrl) {
  const allUrls = scenes.map(function(s) { return s.clipUrl; })
    .concat(scenes.map(function(s) { return s.sceneAudioUrl || ""; }).filter(Boolean))
    .concat(audioUrl ? [audioUrl] : []);
  allUrls.forEach(function(url) {
    if (!url || url === "") return;
    try {
      const fileId = extractDriveFileId(url);
      if (!fileId) return;
      const file = DriveApp.getFileById(fileId);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log("Made public: " + fileId);
    } catch (e) {
      Logger.log("Could not make public (non-fatal): " + e.message);
    }
  });
}


// ══════════════════════════════════════════════════════════════════════════════
// UPLOAD CLIP TO SHOTSTACK (legacy direct upload helper)
// Downloads a Drive file and uploads binary directly to Shotstack.
// Used as a fallback when the two-step init/PUT flow is unavailable.
// ══════════════════════════════════════════════════════════════════════════════
function uploadClipToShotstack(clipUrl, apiKey) {
  try {
    const fileId = extractDriveFileId(clipUrl);
    if (!fileId) return null;

    const file     = DriveApp.getFileById(fileId);
    const blob     = file.getBlob();
    const mimeType = blob.getContentType() || "video/mp4";
    const fileName = file.getName();

    Logger.log("Uploading to Shotstack: " + fileName + " (" + blob.getBytes().length + " bytes)");

    const uploadResponse = UrlFetchApp.fetch("https://api.shotstack.io/ingest/v1/upload", {
      method           : "post",
      contentType      : mimeType,
      headers          : {
        "x-api-key"       : apiKey,
        "Content-Type"    : mimeType,
        "x-shotstack-name": encodeURIComponent(fileName)
      },
      payload          : blob.getBytes(),
      muteHttpExceptions: true
    });

    const code = uploadResponse.getResponseCode();
    const body = JSON.parse(uploadResponse.getContentText());
    Logger.log("Shotstack upload response " + code + ": " + JSON.stringify(body).substring(0, 200));

    if ((code === 200 || code === 201) && body.data && body.data.attributes && body.data.attributes.url) {
      return body.data.attributes.url;
    }
    return null;
  } catch (e) {
    Logger.log("Direct upload failed (non-fatal): " + e.message);
    return null;
  }
}

function extractDriveFileId(url) {
  if (!url) return null;
  const m1 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) return m1[1];
  const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m2) return m2[1];
  return null;
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
// ── getOrCreateContentFolder / getOrCreateScenesFolder ───────────────────────
// These helpers are defined in Pipeline.gs (canonical).
// Stage 9B calls them directly from that shared scope.

function timestampToSeconds(ts) {
  if (!ts || ts === "") return 0;
  const parts = ts.toString().split(":");
  if (parts.length === 2) return parseInt(parts[0]) * 60 + parseInt(parts[1]);
  if (parts.length === 3) return parseInt(parts[0]) * 3600 + parseInt(parts[1]) * 60 + parseInt(parts[2]);
  return 0;
}

function getDefaultDuration(type) {
  const durations = {
    "Checkpoint"  : 5,
    "Infographic" : 8,
    "Text"        : 6,
    "Stock"       : 8,
    "AI Generated": 8,
    "Timeline"    : 10,
    "Animation"   : 6,
    "Minimal"     : 5
  };
  return durations[type] || 7;
}

function convertDriveUrl(url) {
  if (!url) return url;
  var fileId = null;
  const m1 = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m1) fileId = m1[1];
  if (!fileId) {
    const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m2) fileId = m2[1];
  }
  if (fileId) {
    return "https://drive.usercontent.google.com/download?id=" + fileId + "&export=download&authuser=0";
  }
  return url;
}

// ── Bug 3 fix: full caption text, no truncation, word-wrap, red border ───────
function buildCheckpointHtml(desc, isArabic) {
  const dir      = isArabic ? "rtl" : "ltr";
  const align    = isArabic ? "right" : "left";
  const font     = isArabic ? "Cairo,Arial,sans-serif" : "Montserrat,Arial,sans-serif";
  const safeDesc = (desc || "").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return "<div style=\"width:1920px;padding:20px 80px;box-sizing:border-box;" +
    "background:rgba(10,10,10,0.85);border-left:6px solid #FF0000;" +
    "direction:" + dir + ";font-family:" + font + ";\">" +
    "<div style=\"font-size:22px;color:#FFFFFF;text-align:" + align + ";line-height:1.6;" +
    "max-width:1760px;word-wrap:break-word;white-space:normal;\">" +
    safeDesc + "</div></div>";
}

function getShotstackTransition(type) {
  return { in: "fade" };
}