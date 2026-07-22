/* ============================================================================
   menu.gs — GovernX Content OS
   Builds the GovernX toolbar menu on sheet open
   Also contains the Error Log writer used across all pipeline stages

   VERSION 3.1 CHANGES:
   + Stage 9C  — Assemble Film (Remotion), one synced MP4 — REPLACES Stage 8D
                 (per-scene render) and Stage 9B (Shotstack); both retired.
   + Preview One Scene — routes REMOTION_DATA through adapt.js (real case-file
                 component), not the legacy /render compositions.
   + Thumbnail — Remotion only (dark / paper / cinematic styles → Thumbnails
                 subfolder); AI-image (Ideogram) item retired.
   + Stage 10  — chapters built from real scene timestamps (run after Stage 9C).
   + Dashboard — pipeline tracker extended to S10/Thumb/S11 + economics.
   ============================================================================ */

function onOpen() {
  const ui   = SpreadsheetApp.getUi();
  const menu = ui.createMenu("🎬 GovernX");

  menu
    // ── Content Pipeline — one linear flow, run top to bottom ─────────────────
    // Research-verification steps (🔎) are interleaved at the exact point they
    // belong so no step gets skipped. They call the same functions the old
    // standalone "🔎 Research" menu used (defined in research_bridge.gs).
    // "Generate ID (manual)" removed — IDs are created automatically: onEdit
    // covers a typed/pasted Company, and getActiveIdeaRow() generates on demand
    // for rows written programmatically (Stage 0), which simple triggers miss.
    .addItem("0️⃣  Stage 0  — Company Selector",            "runCompanySelector")
    .addSeparator()
    .addItem("1️⃣  Stage 1  — Generate Master Content",     "generateMasterContent")
    .addItem("2️⃣  Stage 2  — Generate Research Database",  "generateResearchDatabase")
    .addItem("🔎 Research ② Run verified research…",         "runVerifiedResearch")
    .addItem("🔎 Research ②b Suggest which claims to use",   "suggestClaims")
    .addItem("🔎 Research ③ Approve ticked claims → Data Moments", "approveSelectedClaims")
    .addItem("🔎 Research ⑤ Publish → Research Database (feeds Stage 3)", "publishToResearchDatabase")
    .addItem("3️⃣  Stage 3  — Generate Script",             "generateScript")
    .addItem("3️⃣B Stage 3B — QA Review (run after Stage 3)", "reviewScript")
    .addItem("4️⃣  Stage 4  — Generate Scenes",              "generateScenes")
    .addItem("4️⃣B Stage 4B — Director Review (run after 4)", "runDirectorReview")
    .addItem("🔎 Research ⑥ Validate scene numbers (run after 4B)", "validateSceneNumbers")
    .addItem("5️⃣  Stage 5  — Create Publishing Row",       "createPublishingRow")
    .addItem("6️⃣  Stage 6  — Export Production Package",   "exportProductionPackage")
    .addItem("7️⃣  Stage 7  — Generate Voiceover Audio",    "generateVoiceover")
    .addItem("7️⃣B Stage 7B — Generate Scene Voiceovers",   "generateSceneVoiceovers")
    .addItem("🔎 Research ④ Validate before render (blocks)", "validateBeforeRender")
    .addSeparator()

    // ── Assembly & Publishing ─────────────────────────────────────────────────
    // RETIRED from the menu (functions kept in their .gs files, nothing deleted):
    //   • Stage 8D "All Scenes"  (renderRemotionScenes)  — rendered every scene to
    //     its own MP4 purely so Shotstack could stitch them. "Assemble Film"
    //     renders the whole film in ONE pass straight from REMOTION_DATA and never
    //     reads AI_CLIP_URL, so those clips now cost ~20 min and are used by nothing.
    //   • Stage 9B  Shotstack     (renderWithShotstack)  — replaced by Assemble Film,
    //     and it was the only thing needing SHOTSTACK_API_KEY.
    // Nothing gates on either (no checkPreviousStage references them) and Stages
    // 10/11 depend on Stage 3/4, not on rendering. To restore, re-add the two
    // .addItem lines — the code is untouched.
    .addItem("🔍  Preview One Scene (check before assembling)",  "previewOneScene")
    .addItem("🎬  Assemble Film (Remotion)",                     "assembleFilmRemotion")
    .addItem("🎬  Check Assembly Status",                        "checkAssembleFilmRemotion")
    .addItem("🔟  Stage 10 — Generate YouTube Metadata",         "generateYouTubeMetadata")
    .addItem("1️⃣1️⃣ Stage 11 — Upload to YouTube (Private)",      "uploadToYouTube")
    .addItem("🚀  Publish Video (Make Public)",                   "publishYouTubeVideo")
    .addItem("🗓️  Schedule Publish (Public later)",              "schedulePublishYouTubeVideo")
    .addSeparator()

    // ── Tools ─────────────────────────────────────────────────────────────────
    // "Cleanup Scenes Folder" (cleanupScenesFolder) is RETIRED from the menu —
    // and this one was actively dangerous once Stage 8D went. It empties the
    // content's Scenes subfolder, which used to hold 8D's disposable scene MP4s.
    // Stage 7B writes the per-scene narration MP3s to that same folder and
    // Assemble Film downloads them by Drive id, so the button's only remaining
    // effect was deleting the audio the assembly depends on — every scene would
    // render silent and the col-23 URLs would 404. Function left in pipeline.gs.
    // Thumbnail is rendered in Remotion, not generated by an image model: the
    // headline is TEXT, and diffusion models garble it (a reference poster came
    // back reading "SCCAN DAL"). Same components + same verified figure as the
    // film, so the two can never disagree. The AI-image (Ideogram) item was
    // retired — it produced fabricated data and trademarked logos, off-brand for
    // an evidence channel. For a photographic look, generate a background image
    // with NO text and composite it via `thumb_photo=` in the Idea Catalogue Note.
    .addItem("🖼️  Generate Thumbnail (Remotion)",                "generateThumbnailRemotion")
    .addItem("📊  Build Dashboard",                              "buildDashboard")
    .addSeparator()
    .addItem("💰  Build / Refresh Video Economics",              "refreshVideoEconomics")
    .addItem("📈  Pull Video Revenue (YouTube Analytics)",       "pullEconomicsRevenue")
    .addItem("🩺  Check render engine is reachable",             "pingResearchEngine")
    .addSeparator()
    // One-time setup — needed once per spreadsheet (both are find-or-create, so
    // they are safe to re-run, but they never need pressing in normal use).
    .addSubMenu(ui.createMenu("⚙️  Setup (one-time)")
      .addItem("Setup research tabs",                            "setupResearchBridge")
      .addItem("Add research columns to Idea Catalogue",         "addResearchColumnsToIdeaCatalogue"))
    .addSeparator()
    .addItem("📋  View Error Log",                               "openErrorLog")
    .addItem("🗑️   Clear Error Log",                             "clearErrorLog")
    .addSeparator()
    .addItem("ℹ️   About GovernX Content OS",                    "showAbout")
    .addToUi();

  // ── Additional menus ────────────────────────────────────────────────────────
  // Each is wrapped: a failure in one must not stop the others from being built.
  // (Menus created by running a builder manually in the editor last only for
  //  that session — they must be called from onOpen to survive a reload.)
  try { buildIntelligenceMenu(); }
  catch (e) { Logger.log("Intelligence menu failed to build: " + e.message); }

  // The standalone "🔎 Research" menu is now merged into the main GovernX menu
  // above (research steps interleaved at their pipeline positions), so it is no
  // longer built separately. buildResearchMenu() still exists in
  // research_bridge.gs if you ever want the separate menu back.
  // try { buildResearchMenu(); }
  // catch (e) { Logger.log("Research menu failed to build: " + e.message); }
}


// ── Error Log writer — called from all pipeline stages ───────────────────────
function logError(stage, contentId, errorType, details) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let   sheet = ss.getSheetByName(SHEET.ERROR_LOG);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET.ERROR_LOG);
      sheet.getRange(1, COL_ERROR.TIMESTAMP ).setValue("Timestamp");
      sheet.getRange(1, COL_ERROR.STAGE     ).setValue("Stage");
      sheet.getRange(1, COL_ERROR.ID        ).setValue("Content ID");
      sheet.getRange(1, COL_ERROR.ERROR_TYPE).setValue("Error Type");
      sheet.getRange(1, COL_ERROR.DETAILS   ).setValue("Details");
      sheet.getRange(1, COL_ERROR.RESOLVED  ).setValue("Resolved?");
      sheet.getRange(1, 1, 1, 6)
           .setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setColumnWidth(COL_ERROR.TIMESTAMP,  160);
      sheet.setColumnWidth(COL_ERROR.STAGE,      200);
      sheet.setColumnWidth(COL_ERROR.ID,         160);
      sheet.setColumnWidth(COL_ERROR.ERROR_TYPE, 180);
      sheet.setColumnWidth(COL_ERROR.DETAILS,    400);
      sheet.setColumnWidth(COL_ERROR.RESOLVED,   100);
    }

    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, COL_ERROR.TIMESTAMP ).setValue(new Date());
    sheet.getRange(lastRow, COL_ERROR.STAGE     ).setValue(stage);
    sheet.getRange(lastRow, COL_ERROR.ID        ).setValue(contentId || "—");
    sheet.getRange(lastRow, COL_ERROR.ERROR_TYPE).setValue(errorType);
    sheet.getRange(lastRow, COL_ERROR.DETAILS   ).setValue(details);
    sheet.getRange(lastRow, COL_ERROR.RESOLVED  ).insertCheckboxes();
    sheet.getRange(lastRow, 1, 1, 6).setBackground("#ffe0e0");

    Logger.log("ERROR LOGGED — " + stage + " | " + contentId + " | " + errorType + ": " + details);

  } catch (e) {
    Logger.log("CRITICAL: Error logging failed: " + e.message);
    Logger.log("Original error: " + stage + " | " + errorType + " | " + details);
  }
}


// ── Open Error Log tab ────────────────────────────────────────────────────────
function openErrorLog() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ERROR_LOG);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      "No errors logged yet. The Error Log tab is created automatically on first error."
    );
    return;
  }
  ss.setActiveSheet(sheet);
}


// ── Clear Error Log ───────────────────────────────────────────────────────────
function clearErrorLog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "Clear Error Log",
    "Clear all error log entries? This cannot be undone.",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.ERROR_LOG);
  if (!sheet) { ui.alert("Error Log tab does not exist yet."); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  ui.alert("✅ Error Log cleared.");
}


// ── About ─────────────────────────────────────────────────────────────────────
function showAbout() {
  SpreadsheetApp.getUi().alert(
    "GovernX Content OS",
    "Version 3.1\n\n" +
    "AI-powered, data-driven content pipeline for the GovernX YouTube channel.\n" +
    "No stock footage. No AI video clips. Verified figures, rendered in Remotion.\n\n" +

    "PIPELINE:\n" +
    "Stage 0   → Company Selector (scores ideas vs Channel Memory)\n" +
    "Stage 1   → Master Content Table (brief auto-generated)\n" +
    "Stage 2   → Research → verified-claims engine (②/②b suggest, ③ build, ⑤ validate)\n" +
    "Stage 3   → Script Bank (+ 3B QA review)\n" +
    "Stage 4   → Visual Library + Director, two-pass (+ 4B review)\n" +
    "⑥         → Validate scene numbers — evidence gate (every figure must trace to a claim)\n" +
    "Stage 5   → Publishing Tracker\n" +
    "Stage 6   → Production Package (Google Doc)\n" +
    "Stage 7   → Voiceover — full-script MP3 (ElevenLabs)\n" +
    "Stage 7B  → Scene Voiceovers — per-scene MP3 (numbers spoken naturally)\n" +
    "④         → Validate before render (blocks)\n" +
    "🔍 Preview One Scene — render one scene as the real case-file component\n" +
    "Stage 9C  → Assemble Film (Remotion) — one synced MP4 (replaces old 8D + 9B Shotstack)\n" +
    "Stage 10  → YouTube Metadata (run AFTER assembly, so chapters are real)\n" +
    "Stage 11  → YouTube Upload — Private or Scheduled; sets the thumbnail\n" +
    "🚀 Publish Video — flip Private → Public\n\n" +

    "TOOLS:\n" +
    "🖼️ Generate Thumbnail (Remotion) — 1280×720 case-file poster; pick\n" +
    "   dark / paper / cinematic; saved to the content's Thumbnails subfolder.\n" +
    "📊 Build Dashboard — per-video pipeline status + portfolio economics.\n" +
    "💰 Video Economics — per-video cost vs revenue, ROI.\n" +
    "🩺 Check render engine is reachable.\n\n" +

    "ID FORMAT: GX-YYMM-DOMAIN-SEQ  (e.g. GX-2607-BIZ-001)\n\n" +
    "Built for GovernX — reverse-engineering why systems fail, and why they win.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}