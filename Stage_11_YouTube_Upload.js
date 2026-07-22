/* ============================================================================
   Stage_11_YouTube_Upload.gs — GovernX Content OS
   Stage 11 — Automated YouTube Upload
   ============================================================================ */

// ══════════════════════════════════════════════════════════════════════════════
// STAGE 11 — MAIN UPLOAD FUNCTION
// ══════════════════════════════════════════════════════════════════════════════
function uploadToYouTube(idea) {

  if (!idea) {
    idea = getActiveIdeaRow();
    if (!idea) return;
  }

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Step 1: Read final video URL from Publishing Tracker ──────────────────
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) {
    ui.alert("⚠️ Stage 11 Failed", "Publishing Tracker tab not found.", ui.ButtonSet.OK);
    return;
  }

  let finalVideoUrl = "";
  let pubRow        = -1;
  const pubData     = pubSheet.getDataRange().getValues();

  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== idea.id) continue;
    pubRow = i + 1;
    const rawUrl = pubData[i][COL_PUBLISHING.SCENES_FOLDER - 1].toString().trim();
    if (rawUrl.startsWith("http")) {
      finalVideoUrl = rawUrl;
    } else {
      const formula = pubSheet.getRange(i + 1, COL_PUBLISHING.SCENES_FOLDER).getFormula();
      const urlMatch = formula.match(/HYPERLINK\("(https?:[^"]+)"/i);
      if (urlMatch) finalVideoUrl = urlMatch[1];
    }
    break;
  }

  if (!finalVideoUrl) {
    ui.alert(
      "⚠️ Stage 11 — No Video Found",
      "No final video URL found in Publishing Tracker for: " + idea.id + "\n\n" +
      "Run 🎬 Assemble Film (Remotion) first to produce the final MP4.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Check if already uploaded
  const existingVideoId = pubData[pubRow - 1] && pubData[pubRow - 1][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1]
    ? pubData[pubRow - 1][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1].toString().trim()
    : "";

  if (existingVideoId && existingVideoId.length > 5) {
    const reupload = ui.alert(
      "Video Already Uploaded",
      "This video was already uploaded to YouTube.\nVideo ID: " + existingVideoId + "\n\n" +
      "Re-upload as a new video?",
      ui.ButtonSet.YES_NO
    );
    if (reupload !== ui.Button.YES) return;
  }

  // ── Step 2: Read metadata from YouTube Metadata tab ──────────────────────
  const ytMeta = readYouTubeMetadata_(ss, idea.id);
  if (!ytMeta) {
    ui.alert(
      "⚠️ Stage 11 — No Metadata Found",
      "YouTube Metadata tab has no entry for: " + idea.id + "\n\nRun Stage 10 first.",
      ui.ButtonSet.OK
    );
    return;
  }

  // ── Step 3: Build full description ───────────────────────────────────────
  const fullDescription = buildFullDescription_(ss, idea, ytMeta);

  // ── Step 4: Extract Drive file ID ────────────────────────────────────────
  const fileId = extractDriveId_(finalVideoUrl);
  if (!fileId) {
    ui.alert(
      "⚠️ Stage 11 Failed",
      "Could not extract Drive file ID from URL:\n" + finalVideoUrl,
      ui.ButtonSet.OK
    );
    return;
  }

  // ── PRE-FLIGHT: sanitize ALL metadata fields ──────────────────────────────
  // \x20-\x7E strips control characters (tabs, newlines, null bytes) that cause
  // YouTube API to reject tags with "invalid video keywords" error
  const sanitize = (str) => (str || "")
    .replace(/[\x00-\x1F\x7F-\x9F]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  const cleanTitle = sanitize(ytMeta.titleA).substring(0, 100) ||
                     sanitize(ytMeta.titleB).substring(0, 100) ||
                     idea.company.substring(0, 100);

  const cleanDescription = sanitize(fullDescription).substring(0, 5000);

  const tagsArray = (ytMeta.tags || "")
    .split(",")
    .map(t => sanitize(t).replace(/[<>]/g, "").replace(/^#+/, "").substring(0, 100))
    .filter(t => t.length >= 2)
    .reduce((acc, tag) => {
      const totalSoFar = acc.length === 0 ? 0 : acc.join(",").length + 1;
      if (totalSoFar + tag.length <= 500) acc.push(tag);
      return acc;
    }, []);

  Logger.log("=== Stage 11 PRE-FLIGHT for " + idea.id + " ===");
  Logger.log("titleA clean    : " + cleanTitle);
  Logger.log("tags count      : " + tagsArray.length);
  Logger.log("tags joined len : " + tagsArray.join(",").length);
  Logger.log("tags            : " + JSON.stringify(tagsArray));
  Logger.log("description len : " + cleanDescription.length);
  Logger.log("=== END PRE-FLIGHT ===");

  // ── Choose upload mode: Private, Schedule, or Public ────────────────────
  const modeResponse = ui.alert(
    "🎬 Stage 11 — Upload to YouTube",
    "Ready to upload:\n\n" +
    "Title    : " + cleanTitle + "\n" +
    "Tags     : " + tagsArray.length + " tags (" + tagsArray.join(",").length + " chars)\n\n" +
    "Choose upload mode:\n\n" +
    "YES  → Upload as PRIVATE (review in YouTube Studio before publishing)\n" +
    "NO   → Schedule for future publish (you will enter a date/time)\n\n" +
    "Upload will take 1–3 minutes.",
    ui.ButtonSet.YES_NO
  );
  if (modeResponse === ui.Button.CLOSE) return;

  let publishAt    = null;  // ISO 8601 string, e.g. "2026-06-10T18:00:00Z"
  let uploadMode   = "private";

  if (modeResponse === ui.Button.NO) {
    // ── Schedule mode: prompt for date and time ──────────────────────────────
    const datePrompt = ui.prompt(
      "📅 Schedule Publish Date",
      "Enter the publish date and time in this format:\n\n" +
      "YYYY-MM-DD HH:MM  (24-hour, your local time)\n\n" +
      "Examples:\n" +
      "  2026-06-15 18:00\n" +
      "  2026-07-01 09:30\n\n" +
      "The video will go public automatically at this time.",
      ui.ButtonSet.OK_CANCEL
    );
    if (datePrompt.getSelectedButton() !== ui.Button.OK) return;

    const rawDate = datePrompt.getResponseText().trim();
    publishAt = parseScheduleDate_(rawDate);

    if (!publishAt) {
      ui.alert(
        "⚠️ Invalid Date Format",
        "Could not parse: '" + rawDate + "'\n\n" +
        "Please use: YYYY-MM-DD HH:MM (e.g. 2026-06-15 18:00)\n\n" +
        "Upload cancelled.",
        ui.ButtonSet.OK
      );
      return;
    }

    // Must be at least 10 minutes in the future
    const publishDate = new Date(publishAt);
    if (publishDate.getTime() < Date.now() + 10 * 60 * 1000) {
      ui.alert(
        "⚠️ Date Too Soon",
        "Scheduled publish time must be at least 10 minutes in the future.\n\n" +
        "Parsed: " + publishDate.toUTCString() + "\n\nUpload cancelled.",
        ui.ButtonSet.OK
      );
      return;
    }

    uploadMode = "scheduled";

    const scheduleConfirm = ui.alert(
      "📅 Confirm Schedule",
      "Video will be uploaded as PRIVATE and auto-published at:\n\n" +
      "🕐 " + publishDate.toUTCString() + "\n\n" +
      "Title: " + cleanTitle + "\n\nProceed?",
      ui.ButtonSet.YES_NO
    );
    if (scheduleConfirm !== ui.Button.YES) return;
  }

  ui.alert(
    "⏳ Uploading to YouTube...",
    "Uploading: " + idea.company + "\n\nThis will take 1–3 minutes.\nDo not close the sheet.\nClick OK to start.",
    ui.ButtonSet.OK
  );

  try {
    // ── Step 5: Load video blob from Drive ───────────────────────────────────
    const videoFile = DriveApp.getFileById(fileId);
    const videoBlob = videoFile.getBlob().setContentType("video/mp4");

    Logger.log("Stage 11: Uploading " + videoFile.getName() +
      " (" + Math.round(videoBlob.getBytes().length / 1024 / 1024) + " MB)" +
      " | mode: " + uploadMode + (publishAt ? " | publishAt: " + publishAt : ""));

    // ── Step 6: Upload via YouTube Data API v3 ───────────────────────────────
    const statusBlock = {
      selfDeclaredMadeForKids: false
    };

    if (uploadMode === "scheduled" && publishAt) {
      // Scheduled: must be private + publishAt for YouTube to auto-publish
      statusBlock.privacyStatus = "private";
      statusBlock.publishAt     = publishAt;
    } else {
      statusBlock.privacyStatus = "private";
    }

    const videoResource = {
      snippet: {
        title      : cleanTitle,
        description: cleanDescription,
        tags       : tagsArray,
        categoryId : "27"
      },
      status: statusBlock
    };

    Logger.log("Stage 11: Sending to YouTube API...");

    const uploadResponse = YouTube.Videos.insert(
      videoResource,
      "snippet,status",
      videoBlob
    );

    const videoId  = uploadResponse.id;
    const videoUrl = "https://www.youtube.com/watch?v=" + videoId;

    Logger.log("Stage 11: Upload successful. Video ID: " + videoId);

    // ── Step 6b: Set the custom thumbnail ────────────────────────────────────
    // Nothing ever set one before: Stage 8E produced an image, wrote it to the
    // Publishing Tracker, and the automation stopped there — every upload needed
    // the thumbnail attached by hand. Non-fatal by design: a failure here must
    // not cost an otherwise successful upload.
    try {
      const thumbUrl = readPublishingThumbnailUrl_(pubSheet, pubRow);
      if (thumbUrl) {
        const idMatch = thumbUrl.match(/\/d\/([a-zA-Z0-9_-]+)/) || thumbUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
        if (idMatch) {
          const blob = DriveApp.getFileById(idMatch[1]).getBlob();
          if (blob.getBytes().length > 2 * 1024 * 1024) {      // YouTube's limit
            Logger.log("Stage 11: thumbnail skipped — over YouTube's 2MB limit.");
          } else {
            YouTube.Thumbnails.set(videoId, blob);
            Logger.log("Stage 11: thumbnail set from Drive id " + idMatch[1]);
          }
        } else {
          Logger.log("Stage 11: thumbnail skipped — no Drive id in " + thumbUrl);
        }
      } else {
        Logger.log("Stage 11: no thumbnail linked — run 🖼️ Generate Thumbnail first.");
      }
    } catch (thumbErr) {
      Logger.log("Stage 11: thumbnail set failed (non-fatal): " + thumbErr.message);
    }

    // ── Step 7: Post pinned first comment ────────────────────────────────────
    if (ytMeta.firstComment && ytMeta.firstComment.trim() !== "") {
      try {
        YouTube.CommentThreads.insert(
          {
            snippet: {
              videoId        : videoId,
              topLevelComment: {
                snippet: { textOriginal: ytMeta.firstComment.substring(0, 10000) }
              }
            }
          },
          "snippet"
        );
        Logger.log("Stage 11: Pinned comment posted.");
      } catch (commentErr) {
        Logger.log("Stage 11: Comment post failed (non-fatal): " + commentErr.message);
      }
    }

    // ── Step 8: Write videoId back to Publishing Tracker ─────────────────────
    if (pubRow > 0) {
      pubSheet.getRange(pubRow, COL_PUBLISHING.YOUTUBE_VIDEO_ID).setValue(videoId);
      SpreadsheetApp.flush();
    }

    // ── Step 9: Update YouTube Metadata tab STATUS ────────────────────────────
    const metaStatus = uploadMode === "scheduled"
      ? "Scheduled — " + new Date(publishAt).toUTCString()
      : "Uploaded (Private)";
    updateYouTubeMetadataStatus_(ss, idea.id, metaStatus);

    // ── Step 10: Write channel memory ─────────────────────────────────────────
    try { writeChannelMemoryForSelected(); } catch (e) {
      Logger.log("Stage 11: Channel memory write skipped: " + e.message);
    }

    updatePipelineStatus_(idea.id, "S11", "✅");

    const successMsg = uploadMode === "scheduled"
      ? "Video uploaded and scheduled for auto-publish:\n\n" +
        "📺 " + videoUrl + "\n\n" +
        "Video ID: " + videoId + "\n" +
        "Title: " + cleanTitle + "\n\n" +
        "🕐 Scheduled publish: " + new Date(publishAt).toUTCString() + "\n\n" +
        "Pinned comment: " + (ytMeta.firstComment ? "✅ Posted" : "— not set") + "\n\n" +
        "The video will go public automatically at the scheduled time.\n" +
        "To publish early: GovernX menu → 🚀 Publish Video (Make Public)"
      : "Video uploaded successfully as PRIVATE:\n\n" +
        "📺 " + videoUrl + "\n\n" +
        "Video ID: " + videoId + "\n" +
        "Title: " + cleanTitle + "\n\n" +
        "Pinned comment: " + (ytMeta.firstComment ? "✅ Posted" : "— not set") + "\n\n" +
        "To publish:\n" +
        "1. Open YouTube Studio → find the video (Private)\n" +
        "2. Change visibility to Public\n\n" +
        "Or use GovernX menu → 🚀 Publish Video (Make Public)";

    ui.alert(
      "✅ Stage 11 Complete — Uploaded to YouTube",
      successMsg,
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S11", "❌");
    Logger.log("Stage 11 ERROR: " + err.message);
    logError("Stage 11 — YouTube Upload", idea.id, "Upload Error", err.message);
    ui.alert(
      "❌ Stage 11 Failed",
      err.message + "\n\nSee Error Log tab.",
      ui.ButtonSet.OK
    );
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// PUBLISH VIDEO — Make a Private video Public
// ══════════════════════════════════════════════════════════════════════════════
function publishYouTubeVideo() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui       = SpreadsheetApp.getUi();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) { ui.alert("Publishing Tracker not found."); return; }

  let videoId = "";
  const pubData = pubSheet.getDataRange().getValues();
  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== idea.id) continue;
    videoId = pubData[i][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1].toString().trim();
    break;
  }

  if (!videoId || videoId.length < 5) {
    ui.alert("No Video ID Found",
      "No YouTube Video ID in Publishing Tracker for: " + idea.id + "\n\nRun Stage 11 first.",
      ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    "🚀 Publish Video",
    "Make this video PUBLIC on YouTube?\n\n" +
    "Video ID: " + videoId + "\n" +
    "URL: https://www.youtube.com/watch?v=" + videoId,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    YouTube.Videos.update(
      { id: videoId, status: { privacyStatus: "public" } },
      "status"
    );
    updateYouTubeMetadataStatus_(ss, idea.id, "Published ✅");
    for (let i = 1; i < pubData.length; i++) {
      if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== idea.id) continue;
      pubSheet.getRange(i + 1, COL_PUBLISHING.PUBLISH_DATE).setValue(new Date());
      break;
    }
    SpreadsheetApp.flush();
    ui.alert("✅ Video Published",
      "Your video is now PUBLIC:\n\nhttps://www.youtube.com/watch?v=" + videoId,
      ui.ButtonSet.OK);
  } catch (err) {
    logError("Stage 11 — Publish", idea.id, "Publish Error", err.message);
    ui.alert("❌ Publish Failed", err.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// SCHEDULE PUBLISH — set a future publish time on an already-uploaded PRIVATE video
// The video stays private until `publishAt`, then YouTube makes it public itself.
// This does NOT upload — it targets the video already uploaded (id in Publishing
// Tracker col R). Use "🚀 Publish Video (Make Public)" to go live before that time.
// ══════════════════════════════════════════════════════════════════════════════
function schedulePublishYouTubeVideo() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui       = SpreadsheetApp.getUi();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) { ui.alert("Publishing Tracker not found."); return; }

  let videoId = "", pubRow = -1;
  const pubData = pubSheet.getDataRange().getValues();
  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== idea.id) continue;
    videoId = pubData[i][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1].toString().trim();
    pubRow  = i + 1;
    break;
  }
  if (!videoId || videoId.length < 5) {
    ui.alert("No Video ID Found",
      "No YouTube Video ID in Publishing Tracker for: " + idea.id + "\n\n" +
      "Upload it first (Stage 11), or paste the video ID into the Publishing Tracker (col R).",
      ui.ButtonSet.OK);
    return;
  }

  // ── Verify the video exists and is schedulable (not already public) ─────────
  let current = null;
  try { current = YouTube.Videos.list("status", { id: videoId }); }
  catch (e) {
    ui.alert("Could not reach YouTube", "Error checking the video's status:\n" + e.message, ui.ButtonSet.OK);
    return;
  }
  if (!current || !current.items || !current.items.length) {
    ui.alert("Video Not Found",
      "No video with ID '" + videoId + "' was found on your channel.\n\n" +
      "Check the ID in the Publishing Tracker (col R).",
      ui.ButtonSet.OK);
    return;
  }
  const curPrivacy = current.items[0].status ? current.items[0].status.privacyStatus : "";
  if (curPrivacy === "public") {
    ui.alert("Already Public",
      "This video is already PUBLIC, so it can't be scheduled.\n\n" +
      "To schedule it, set it back to Private in YouTube Studio first, then run this again.",
      ui.ButtonSet.OK);
    return;
  }
  const curPublishAt = current.items[0].status ? current.items[0].status.publishAt : "";

  // ── Prompt for the publish date/time (same parser Stage 11 uses) ────────────
  const datePrompt = ui.prompt(
    "🗓️ Schedule Publish — " + idea.id,
    "The video stays PRIVATE until this time, then YouTube makes it PUBLIC automatically.\n\n" +
    "Enter the publish date and time:\n" +
    "YYYY-MM-DD HH:MM  (24-hour, your local time)\n\n" +
    "Examples:  2026-07-25 18:00   /   2026-08-01 09:30",
    ui.ButtonSet.OK_CANCEL
  );
  if (datePrompt.getSelectedButton() !== ui.Button.OK) return;

  const publishAt = parseScheduleDate_(datePrompt.getResponseText().trim());
  if (!publishAt) {
    ui.alert("⚠️ Invalid Date Format",
      "Use: YYYY-MM-DD HH:MM (e.g. 2026-07-25 18:00). Cancelled.", ui.ButtonSet.OK);
    return;
  }
  const when = new Date(publishAt);
  if (when.getTime() < Date.now() + 10 * 60 * 1000) {
    ui.alert("⚠️ Date Too Soon",
      "Scheduled time must be at least 10 minutes in the future.\n\nParsed: " +
      when.toUTCString() + "\n\nCancelled.", ui.ButtonSet.OK);
    return;
  }

  const reschedule = curPublishAt
    ? "Currently scheduled: " + new Date(curPublishAt).toUTCString() + " (will be replaced)\n"
    : "";
  const confirm = ui.alert("🗓️ Schedule Publish",
    "Schedule this video to go PUBLIC automatically?\n\n" +
    "Video ID: " + videoId + "\n" +
    reschedule +
    "New publish time: " + when.toUTCString() + "\n" +
    "(stays Private until then)",
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  try {
    // YouTube schedules a video when privacyStatus=private AND publishAt is set (UTC ISO 8601).
    YouTube.Videos.update(
      { id: videoId, status: { privacyStatus: "private", publishAt: publishAt } },
      "status"
    );
    updateYouTubeMetadataStatus_(ss, idea.id, "Scheduled 🕐 " + when.toUTCString());
    if (pubRow > 0) {
      pubSheet.getRange(pubRow, COL_PUBLISHING.NOTES).setValue("Scheduled publish: " + when.toUTCString());
    }
    SpreadsheetApp.flush();
    ui.alert("✅ Publish Scheduled",
      "The video will go PUBLIC automatically at:\n\n" + when.toUTCString() + "\n\n" +
      "https://www.youtube.com/watch?v=" + videoId + "\n\n" +
      "To publish earlier, use 🚀 Publish Video (Make Public).",
      ui.ButtonSet.OK);
  } catch (err) {
    logError("Stage 11 — Schedule Publish", idea.id, "Schedule Error", err.message);
    ui.alert("❌ Schedule Failed", err.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// BUILD FULL DESCRIPTION
// ══════════════════════════════════════════════════════════════════════════════
function buildFullDescription_(ss, idea, ytMeta) {

  let description = ytMeta.description || "";

  const researchSheet = ss.getSheetByName(SHEET.RESEARCH);
  const sources = [];

  if (researchSheet) {
    const resData = researchSheet.getDataRange().getValues();
    for (let i = 1; i < resData.length; i++) {
      if (resData[i][COL_RESEARCH.ID - 1].toString().trim() !== idea.id) continue;
      const sourceType = resData[i][COL_RESEARCH.SOURCE_TYPE - 1].toString().trim();
      if (sourceType === "Data") continue;
      const details   = resData[i][COL_RESEARCH.DETAILS    - 1].toString().trim();
      const sourceUrl = resData[i][COL_RESEARCH.SOURCE_LINK - 1].toString().trim();
      if (!details) continue;
      sources.push({ details, sourceUrl });
    }
  }

  if (sources.length > 0) {
    // Group by URL so a single primary source (e.g. one SEC release cited 15×)
    // is listed ONCE with a few key quotes — not repeated into a wall. Cap the
    // number of distinct sources and the quotes per source so the description
    // stays readable and well under YouTube's 5000-char limit.
    const MAX_SOURCES       = 6;
    const MAX_QUOTES_PER_SRC = 3;
    const byUrl = {}, order = [];
    sources.forEach(function(src) {
      const key = (src.sourceUrl && src.sourceUrl.indexOf("http") === 0) ? src.sourceUrl : ("_nolink_" + order.length);
      if (!byUrl[key]) { byUrl[key] = { url: src.sourceUrl, quotes: [] }; order.push(key); }
      const q = String(src.details || "").trim();
      if (q && byUrl[key].quotes.indexOf(q) === -1 && byUrl[key].quotes.length < MAX_QUOTES_PER_SRC) {
        byUrl[key].quotes.push(q);
      }
    });

    description += "\n\n---\nSOURCES & REFERENCES\n---\n";
    order.slice(0, MAX_SOURCES).forEach(function(key, i) {
      const s = byUrl[key];
      description += "\n[" + (i + 1) + "] ";
      if (s.url && s.url.indexOf("http") === 0) description += s.url;
      s.quotes.forEach(function(q) { description += "\n    • " + q; });
    });
  }

  if (ytMeta.hashtags && ytMeta.hashtags.trim() !== "") {
    const cleanHashtags = (ytMeta.hashtags || "").replace(/[^\x20-\x7E]/g, "").trim();
    if (cleanHashtags) description += "\n\n" + cleanHashtags;
  }

  Logger.log("Stage 11: description built — " + description.length + " chars | " + sources.length + " sources");
  return description;
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
function readYouTubeMetadata_(ss, contentId) {
  const ytSheet = ss.getSheetByName(SHEET.YOUTUBE);
  if (!ytSheet) return null;
  const data = ytSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_YOUTUBE.ID - 1].toString().trim() !== contentId) continue;
    return {
      titleA      : data[i][COL_YOUTUBE.TITLE_A        - 1].toString().trim(),
      titleB      : data[i][COL_YOUTUBE.TITLE_B        - 1].toString().trim(),
      titleC      : data[i][COL_YOUTUBE.TITLE_C        - 1].toString().trim(),
      description : data[i][COL_YOUTUBE.DESCRIPTION    - 1].toString().trim(),
      tags        : data[i][COL_YOUTUBE.TAGS           - 1].toString().trim(),
      chapters    : data[i][COL_YOUTUBE.CHAPTERS       - 1].toString().trim(),
      hashtags    : data[i][COL_YOUTUBE.HASHTAGS       - 1].toString().trim(),
      firstComment: data[i][COL_YOUTUBE.FIRST_COMMENT  - 1].toString().trim(),
      endScreen   : data[i][COL_YOUTUBE.END_SCREEN     - 1].toString().trim(),
      thumbBrief  : data[i][COL_YOUTUBE.THUMBNAIL_BRIEF- 1].toString().trim()
    };
  }
  return null;
}

function updateYouTubeMetadataStatus_(ss, contentId, status) {
  const ytSheet = ss.getSheetByName(SHEET.YOUTUBE);
  if (!ytSheet) return;
  const data = ytSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_YOUTUBE.ID - 1].toString().trim() !== contentId) continue;
    ytSheet.getRange(i + 1, COL_YOUTUBE.STATUS).setValue(status);
    SpreadsheetApp.flush();
    return;
  }
}

function extractDriveId_(url) {
  if (!url) return null;
  const m1 = url.match(/\/d\/([a-zA-Z0-9_-]{10,})/);
  if (m1) return m1[1];
  const m2 = url.match(/[?&]id=([a-zA-Z0-9_-]{10,})/);
  if (m2) return m2[1];
  return null;
}


// ── parseScheduleDate_ ────────────────────────────────────────────────────────
// Parses "YYYY-MM-DD HH:MM" (user's local timezone, as GAS session timezone)
// Returns an ISO 8601 UTC string ("2026-06-15T16:00:00Z") or null on failure.
// YouTube requires publishAt in UTC ISO 8601.
function parseScheduleDate_(rawDate) {
  if (!rawDate) return null;

  // Accept "YYYY-MM-DD HH:MM" with flexible separators between date and time
  const match = rawDate.trim().match(
    /^(\d{4})-(\d{2})-(\d{2})[\sT]+(\d{1,2}):(\d{2})$/
  );
  if (!match) return null;

  const year  = parseInt(match[1], 10);
  const month = parseInt(match[2], 10) - 1; // JS months are 0-indexed
  const day   = parseInt(match[3], 10);
  const hour  = parseInt(match[4], 10);
  const min   = parseInt(match[5], 10);

  // Basic range validation
  if (month < 0 || month > 11) return null;
  if (day < 1   || day > 31  ) return null;
  if (hour < 0  || hour > 23 ) return null;
  if (min < 0   || min > 59  ) return null;

  // Construct as local time using the GAS session timezone
  // new Date(y, m, d, h, min) uses the script's timezone (same as spreadsheet)
  const localDate = new Date(year, month, day, hour, min, 0, 0);
  if (isNaN(localDate.getTime())) return null;

  // Return ISO 8601 UTC string — YouTube API expects UTC
  return localDate.toISOString();
}