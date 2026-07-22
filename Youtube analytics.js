/* ============================================================================
   YouTube_Analytics.gs — GovernX Content OS
   Automatic YouTube Analytics Pull + Publishing Tracker Write-back

   WHAT THIS DOES:
   1. Reads all content IDs from Publishing Tracker that have a YouTube Video ID
   2. Calls YouTube Analytics API v2 to pull views, CTR, retention, watch time
   3. Writes data back to Publishing Tracker columns automatically
   4. Calculates GovernX Performance Score per video
   5. Triggers Channel Memory update for any video with new performance data
   6. Runs automatically on a weekly timed trigger (every Sunday)

   REQUIRED SETUP — DO THIS ONCE:
   In Apps Script editor:
   1. Click "Services" (+ icon in left sidebar)
   2. Add "YouTube Data API v3" → save as "YouTube"
   3. Add "YouTube Analytics API" → save as "YouTubeAnalytics"
   4. Run setupAnalyticsTrigger() once from the menu to install the weekly trigger

   NO API KEYS NEEDED — Apps Script uses your Google account OAuth automatically.
   You must be signed in as the YouTube channel owner.

   MENU INTEGRATION:
   Add to Menu.gs in onOpen():
     .addItem("📺 Pull YouTube Analytics",       "pullAllVideoAnalytics")
     .addItem("⏰ Setup Weekly Analytics Trigger","setupAnalyticsTrigger")
     .addItem("🗑️  Remove Analytics Trigger",     "removeAnalyticsTrigger")

   PUBLISHING TRACKER — NEW COLUMN NEEDED:
   Add a "YouTube Video ID" column to your Publishing Tracker sheet.
   After uploading a video to YouTube, paste the video ID (e.g. "dQw4w9WgXcQ")
   into that column. The analytics puller reads this ID to query the API.

   Column to add: COL_PUBLISHING.YOUTUBE_VIDEO_ID
   Add this to config.gs:
     YOUTUBE_VIDEO_ID : 18   (or whatever the next available column is)
   ============================================================================ */


// ── Column index for YouTube Video ID in Publishing Tracker ──────────────────
// Add this to config.gs: YOUTUBE_VIDEO_ID : 18
// Defined here as fallback if not yet in config.gs
const YT_VIDEO_ID_COL = (typeof COL_PUBLISHING !== "undefined" && COL_PUBLISHING.YOUTUBE_VIDEO_ID)
  ? COL_PUBLISHING.YOUTUBE_VIDEO_ID
  : 18;

// ── Analytics pull window ─────────────────────────────────────────────────────
const ANALYTICS_LOOKBACK_DAYS = 90;  // Pull last 90 days of data per video

// ── Performance score weights (GovernX formula) ───────────────────────────────
// Score = (CTR_normalized × 40) + (retention_normalized × 40) + (views_normalized × 20)
// Normalized = metric / benchmark target
const PERF_WEIGHTS = {
  ctr       : 0.40,   // 40% weight — hook effectiveness
  retention : 0.40,   // 40% weight — content quality
  views     : 0.20    // 20% weight — reach
};


// ══════════════════════════════════════════════════════════════════════════════
// PULL ALL VIDEO ANALYTICS
// Main function — reads all content IDs, pulls analytics, writes back to sheet
// Called manually from menu or automatically by weekly trigger
// ══════════════════════════════════════════════════════════════════════════════
function pullAllVideoAnalytics() {

  const ui       = SpreadsheetApp.getUi();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);

  if (!pubSheet) {
    ui.alert("Publishing Tracker tab not found.");
    return;
  }

  const pubData = pubSheet.getDataRange().getValues();

  if (pubData.length < 2) {
    ui.alert("No videos in Publishing Tracker yet.");
    return;
  }

  // ── Check API services are enabled ─────────────────────────────────────────
  try {
    YouTube.Channels.list("id", { mine: true });
  } catch (e) {
    ui.alert(
      "⚠️ YouTube API Not Enabled",
      "You need to enable two services first:\n\n" +
      "1. In Apps Script editor → click '+' next to 'Services'\n" +
      "2. Add 'YouTube Data API v3' (save as YouTube)\n" +
      "3. Add 'YouTube Analytics API' (save as YouTubeAnalytics)\n" +
      "4. Try again.\n\n" +
      "Error: " + e.message,
      ui.ButtonSet.OK
    );
    return;
  }

  // ── Get channel ID once ─────────────────────────────────────────────────────
  let channelId;
  try {
    const channels = YouTube.Channels.list("id", { mine: true });
    if (!channels.items || channels.items.length === 0) {
      ui.alert("No YouTube channel found for your Google account.");
      return;
    }
    channelId = channels.items[0].id;
    Logger.log("Channel ID: " + channelId);
  } catch (e) {
    ui.alert("Could not retrieve channel ID: " + e.message);
    return;
  }

  // ── Collect videos to process ───────────────────────────────────────────────
  const videosToProcess = [];

  for (let i = 1; i < pubData.length; i++) {
    const contentId  = pubData[i][COL_PUBLISHING.ID           - 1].toString().trim();
    const videoId    = pubData[i][YT_VIDEO_ID_COL             - 1].toString().trim();
    const publishDate= pubData[i][COL_PUBLISHING.PUBLISH_DATE - 1];

    if (!contentId || !videoId || videoId === "") continue;
    if (videoId.length < 8) continue; // not a valid YouTube video ID

    videosToProcess.push({
      row         : i + 1,
      contentId   : contentId,
      videoId     : videoId,
      publishDate : publishDate
    });
  }

  if (videosToProcess.length === 0) {
    ui.alert(
      "No Videos to Process",
      "No rows in Publishing Tracker have a YouTube Video ID.\n\n" +
      "After uploading a video to YouTube, paste the Video ID\n" +
      "(the part after youtube.com/watch?v=) into column " + YT_VIDEO_ID_COL + "\n" +
      "of the Publishing Tracker.",
      ui.ButtonSet.OK
    );
    return;
  }

  // ── Confirm before running ──────────────────────────────────────────────────
  const isManual = typeof ui !== "undefined";
  if (isManual) {
    const confirm = ui.alert(
      "📺 Pull YouTube Analytics",
      "Found " + videosToProcess.length + " video(s) with YouTube IDs.\n\n" +
      "Pulling analytics for each...\n" +
      "(This takes a few seconds per video)\n\n" +
      "Proceed?",
      ui.ButtonSet.YES_NO
    );
    if (confirm !== ui.Button.YES) return;
  }

  // ── Process each video ──────────────────────────────────────────────────────
  let updated = 0;
  let failed  = 0;
  const updatedIds = [];

  videosToProcess.forEach(video => {
    try {
      const metrics = pullVideoMetrics(video.videoId, channelId, video.publishDate);
      if (metrics) {
        writeMetricsToPublishing(pubSheet, video.row, metrics);
        updatedIds.push(video.contentId);
        updated++;
        Logger.log("Analytics updated: " + video.contentId + " | Views: " + metrics.views);
      }
      Utilities.sleep(500); // avoid rate limits
    } catch (err) {
      failed++;
      Logger.log("Analytics failed for " + video.contentId + ": " + err.message);
      logError("YouTube Analytics", video.contentId, "API Error", err.message);
    }
  });

  SpreadsheetApp.flush();

  // ── Trigger Channel Memory update for updated videos ───────────────────────
  updatedIds.forEach(contentId => {
    try {
      writeChannelMemory(contentId, contentId);
      Logger.log("Channel Memory refreshed for: " + contentId);
    } catch (e) {
      Logger.log("Memory refresh failed for " + contentId + ": " + e.message);
    }
  });

  // ── Summary ─────────────────────────────────────────────────────────────────
  if (isManual) {
    ui.alert(
      "✅ Analytics Pull Complete",
      "Updated  : " + updated + " video(s)\n" +
      (failed > 0 ? "Failed   : " + failed + " (see Error Log)\n\n" : "\n") +
      "Publishing Tracker updated with:\n" +
      "• Views\n• CTR\n• Audience Retention\n• Avg Watch Time\n" +
      "• GovernX Performance Score\n\n" +
      "Channel Memory also refreshed for updated videos.",
      ui.ButtonSet.OK
    );
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// PULL VIDEO METRICS
// Calls YouTube Analytics API for a single video
// Returns a metrics object or null if no data
// ══════════════════════════════════════════════════════════════════════════════
function pullVideoMetrics(videoId, channelId, publishDate) {

  // ── Date range ──────────────────────────────────────────────────────────────
  const today     = new Date();
  const startDate = publishDate
    ? Utilities.formatDate(new Date(publishDate), "UTC", "yyyy-MM-dd")
    : Utilities.formatDate(
        new Date(today.getTime() - ANALYTICS_LOOKBACK_DAYS * 24 * 60 * 60 * 1000),
        "UTC", "yyyy-MM-dd"
      );
  const endDate = Utilities.formatDate(today, "UTC", "yyyy-MM-dd");

  // ── Core metrics query ──────────────────────────────────────────────────────
  // views, estimatedMinutesWatched, averageViewDuration, averageViewPercentage
  // impressionsClickThroughRate requires the reach report (separate query)
  let coreData = null;
  try {
    const coreResponse = YouTubeAnalytics.Reports.query({
      ids        : "channel==" + channelId,
      startDate  : startDate,
      endDate    : endDate,
      metrics    : "views,estimatedMinutesWatched,averageViewDuration,averageViewPercentage",
      filters    : "video==" + videoId
    });

    if (coreResponse.rows && coreResponse.rows.length > 0) {
      coreData = {
        views              : coreResponse.rows[0][0] || 0,
        estimatedMinutes   : coreResponse.rows[0][1] || 0,
        avgViewDuration    : coreResponse.rows[0][2] || 0,   // seconds
        avgViewPercentage  : coreResponse.rows[0][3] || 0    // 0–100
      };
    }
  } catch (e) {
    Logger.log("Core metrics query failed for " + videoId + ": " + e.message);
  }

  // ── Reach / impressions query (CTR) ────────────────────────────────────────
  let ctrData = null;
  try {
    const reachResponse = YouTubeAnalytics.Reports.query({
      ids        : "channel==" + channelId,
      startDate  : startDate,
      endDate    : endDate,
      metrics    : "impressions,impressionsClickThroughRate",
      filters    : "video==" + videoId
    });

    if (reachResponse.rows && reachResponse.rows.length > 0) {
      ctrData = {
        impressions : reachResponse.rows[0][0] || 0,
        ctr         : parseFloat((reachResponse.rows[0][1] * 100).toFixed(2)) || 0  // convert to %
      };
    }
  } catch (e) {
    // CTR query sometimes requires additional permissions or may not be available
    // for all channels — fail gracefully
    Logger.log("CTR query failed for " + videoId + " (non-fatal): " + e.message);
    ctrData = { impressions: 0, ctr: 0 };
  }

  // Return null if we got nothing
  if (!coreData && !ctrData) return null;

  // ── Format avg watch time as MM:SS ─────────────────────────────────────────
  const avgSeconds  = coreData ? Math.round(coreData.avgViewDuration) : 0;
  const avgMinutes  = Math.floor(avgSeconds / 60);
  const avgSecs     = avgSeconds % 60;
  const avgWatchFmt = avgMinutes + ":" + String(avgSecs).padStart(2, "0");

  // ── Calculate GovernX Performance Score ─────────────────────────────────────
  // Normalized against GovernX benchmarks from config.gs
  const views     = coreData ? coreData.views             : 0;
  const ctr       = ctrData  ? ctrData.ctr                : 0;
  const retention = coreData ? coreData.avgViewPercentage : 0;

  const ctrNorm       = Math.min(ctr       / BENCHMARKS.CTR_SCALE,       1); // benchmark: 6%
  const retentionNorm = Math.min(retention / BENCHMARKS.RETENTION_SCALE,  1); // benchmark: 45%
  const viewsNorm     = Math.min(views     / (BENCHMARKS.VIEWS_KILL * 4), 1); // benchmark: 2000 views

  const perfScore = Math.round(
    (ctrNorm       * PERF_WEIGHTS.ctr       * 100) +
    (retentionNorm * PERF_WEIGHTS.retention  * 100) +
    (viewsNorm     * PERF_WEIGHTS.views      * 100)
  );

  // ── Performance verdict ─────────────────────────────────────────────────────
  // Based on BENCHMARKS from config.gs
  let perfFormula = "Growing";
  if (ctr >= BENCHMARKS.CTR_SCALE && retention >= BENCHMARKS.RETENTION_SCALE) {
    perfFormula = "Scale ✅";
  } else if (views < BENCHMARKS.VIEWS_KILL && retention < 30) {
    perfFormula = "Kill ⛔";
  } else if (ctr < 3 || retention < 30) {
    perfFormula = "Improve 🔧";
  }

  return {
    views        : views,
    impressions  : ctrData  ? ctrData.impressions           : 0,
    ctr          : ctr,
    retention    : parseFloat(retention.toFixed(1)),
    avgWatchTime : avgWatchFmt,
    avgSeconds   : avgSeconds,
    perfScore    : perfScore,
    perfFormula  : perfFormula,
    pulledAt     : new Date().toLocaleString()
  };
}


// ══════════════════════════════════════════════════════════════════════════════
// WRITE METRICS TO PUBLISHING TRACKER
// Writes fetched metrics into the correct columns of the Publishing Tracker row
// ══════════════════════════════════════════════════════════════════════════════
function writeMetricsToPublishing(pubSheet, row, metrics) {

  // Clear validations on all metric columns before writing
  [
    COL_PUBLISHING.VIEWS,
    COL_PUBLISHING.CTR,
    COL_PUBLISHING.RETENTION,
    COL_PUBLISHING.AVG_WATCH_TIME,
    COL_PUBLISHING.PERF_FORMULA,
    COL_PUBLISHING.PERF_SCORE
  ].forEach(col => {
    pubSheet.getRange(row, col).clearDataValidations();
  });

  pubSheet.getRange(row, COL_PUBLISHING.VIEWS         ).setValue(metrics.views);
  pubSheet.getRange(row, COL_PUBLISHING.CTR           ).setValue(metrics.ctr);
  pubSheet.getRange(row, COL_PUBLISHING.RETENTION     ).setValue(metrics.retention);
  pubSheet.getRange(row, COL_PUBLISHING.AVG_WATCH_TIME).setValue(metrics.avgWatchTime);
  pubSheet.getRange(row, COL_PUBLISHING.PERF_FORMULA  ).setValue(metrics.perfFormula);
  pubSheet.getRange(row, COL_PUBLISHING.PERF_SCORE    ).setValue(metrics.perfScore);

  // ── Color code the performance formula cell ─────────────────────────────────
  const formulaCell = pubSheet.getRange(row, COL_PUBLISHING.PERF_FORMULA);
  if (metrics.perfFormula.includes("Scale")) {
    formulaCell.setBackground("#C8F7C5").setFontColor("#155724"); // green
  } else if (metrics.perfFormula.includes("Kill")) {
    formulaCell.setBackground("#FFD7D7").setFontColor("#721C24"); // red
  } else if (metrics.perfFormula.includes("Improve")) {
    formulaCell.setBackground("#FFF3CD").setFontColor("#856404"); // amber
  } else {
    formulaCell.setBackground("#D6EAF8").setFontColor("#1B4F72"); // blue
  }

  // ── Write analytics pull timestamp to Notes column ─────────────────────────
  const currentNotes = pubSheet.getRange(row, COL_PUBLISHING.NOTES).getValue().toString();
  const timestampNote = "\n📊 Analytics pulled: " + metrics.pulledAt;

  // Only append if timestamp not already there
  if (!currentNotes.includes("Analytics pulled:")) {
    pubSheet.getRange(row, COL_PUBLISHING.NOTES)
      .setValue(currentNotes + timestampNote);
  } else {
    // Update the existing timestamp
    const updated = currentNotes.replace(
      /📊 Analytics pulled: .+/,
      "📊 Analytics pulled: " + metrics.pulledAt
    );
    pubSheet.getRange(row, COL_PUBLISHING.NOTES).setValue(updated);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// PULL ANALYTICS FOR SELECTED VIDEO ONLY
// Useful for checking a single video without processing the entire tracker
// ══════════════════════════════════════════════════════════════════════════════
function pullAnalyticsForSelected() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui       = SpreadsheetApp.getUi();
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);

  if (!pubSheet) {
    ui.alert("Publishing Tracker tab not found.");
    return;
  }

  const pubData = pubSheet.getDataRange().getValues();
  let   targetRow = -1;
  let   videoId   = "";

  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
      targetRow = i + 1;
      videoId   = pubData[i][YT_VIDEO_ID_COL - 1].toString().trim();
      break;
    }
  }

  if (targetRow === -1) {
    ui.alert("No Publishing Tracker row found for: " + idea.id + "\nRun Stage 5 first.");
    return;
  }

  if (!videoId || videoId.length < 8) {
    ui.alert(
      "No YouTube Video ID",
      "No YouTube Video ID found for: " + idea.id + "\n\n" +
      "After uploading to YouTube, paste the Video ID into column " +
      YT_VIDEO_ID_COL + " of the Publishing Tracker.\n\n" +
      "The Video ID is the part after 'youtube.com/watch?v=' in the URL.\n" +
      "Example: for youtube.com/watch?v=dQw4w9WgXcQ, the ID is dQw4w9WgXcQ",
      ui.ButtonSet.OK
    );
    return;
  }

  try {
    const channels = YouTube.Channels.list("id", { mine: true });
    const channelId = channels.items[0].id;

    ui.alert(
      "📺 Pulling Analytics",
      "Pulling YouTube analytics for: " + idea.company +
      "\nVideo ID: " + videoId +
      "\n\nThis takes a few seconds...",
      ui.ButtonSet.OK
    );

    const publishDate = pubData[targetRow - 1][COL_PUBLISHING.PUBLISH_DATE - 1];
    const metrics = pullVideoMetrics(videoId, channelId, publishDate);

    if (!metrics) {
      ui.alert(
        "No Data Available",
        "YouTube Analytics returned no data for this video yet.\n\n" +
        "Analytics data can take 24–48 hours to appear after publishing.\n" +
        "Try again tomorrow.",
        ui.ButtonSet.OK
      );
      return;
    }

    writeMetricsToPublishing(pubSheet, targetRow, metrics);
    SpreadsheetApp.flush();

    // Update Channel Memory
    try {
      writeChannelMemory(idea.id, idea.company);
    } catch (e) {
      Logger.log("Memory update after analytics pull failed: " + e.message);
    }

    ui.alert(
      "✅ Analytics Updated — " + idea.id,
      "Views         : " + metrics.views                + "\n" +
      "CTR           : " + metrics.ctr                  + "%\n" +
      "Retention     : " + metrics.retention             + "%\n" +
      "Avg Watch Time: " + metrics.avgWatchTime          + "\n" +
      "Perf Score    : " + metrics.perfScore             + "/100\n" +
      "Verdict       : " + metrics.perfFormula           + "\n\n" +
      "Publishing Tracker and Channel Memory updated.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("YouTube Analytics", idea.id, "Pull Error", err.message);
    ui.alert("❌ Analytics Pull Failed", err.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// SETUP WEEKLY TRIGGER
// Installs a time-based trigger to pull analytics automatically every Sunday
// Run this once from the menu to activate automatic analytics collection
// ══════════════════════════════════════════════════════════════════════════════
function setupAnalyticsTrigger() {

  const ui = SpreadsheetApp.getUi();

  // Check if trigger already exists
  const existing = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "pullAllVideoAnalytics");

  if (existing.length > 0) {
    ui.alert(
      "Trigger Already Active",
      "Weekly analytics trigger is already set up.\n\n" +
      "It runs every Sunday automatically.\n" +
      "To remove it, use '🗑️ Remove Analytics Trigger' from the menu.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Create weekly trigger — every Sunday at 9am
  ScriptApp.newTrigger("pullAllVideoAnalytics")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(9)
    .create();

  ui.alert(
    "✅ Weekly Trigger Active",
    "YouTube Analytics will now pull automatically every Sunday at 9am.\n\n" +
    "Data flows:\n" +
    "YouTube API → Publishing Tracker → Channel Memory → Stage 1 brief\n\n" +
    "You can still pull manually anytime using:\n" +
    "'📺 Pull YouTube Analytics' from the menu.",
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// REMOVE ANALYTICS TRIGGER
// ══════════════════════════════════════════════════════════════════════════════
function removeAnalyticsTrigger() {

  const ui = SpreadsheetApp.getUi();

  const triggers = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "pullAllVideoAnalytics");

  if (triggers.length === 0) {
    ui.alert("No analytics trigger found.", "", ui.ButtonSet.OK);
    return;
  }

  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  ui.alert(
    "✅ Trigger Removed",
    "Weekly analytics trigger has been removed.\n\n" +
    "You can still pull analytics manually from the menu.",
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// CONFIG.GS ADDITION REQUIRED
// Add this line to the COL_PUBLISHING constant in config.gs:
//
//   YOUTUBE_VIDEO_ID : 18   // ← new column — paste YouTube video ID here after upload
//
// And shift NOTES from 17 to 19, or keep NOTES at 17 and add YOUTUBE_VIDEO_ID
// as column 18 in your actual sheet (insert a new column after NOTES).
//
// MENU.GS ADDITIONS REQUIRED
// Add these items to onOpen() after the analytics trigger separator:
//
//   .addSeparator()
//   .addItem("📺 Pull YouTube Analytics (All)",     "pullAllVideoAnalytics")
//   .addItem("📺 Pull Analytics for Selected",      "pullAnalyticsForSelected")
//   .addItem("⏰ Setup Weekly Analytics Trigger",   "setupAnalyticsTrigger")
//   .addItem("🗑️  Remove Analytics Trigger",         "removeAnalyticsTrigger")
// ══════════════════════════════════════════════════════════════════════════════