/* ============================================================================
   Stage_8E_Thumbnail.gs — GovernX Content OS
   Stage 8E — Thumbnail Brief Formatter

   FIX-7: THUMBNAIL_BRIEF is written by Stage 1 into Master Content Table
   but never processed downstream. This stage reads it and outputs:
     A) Structured Canva spec (colors, fonts, layout, dimensions)
     B) AI image generation prompt (Midjourney / DALL-E / Ideogram ready)
     C) Stop-scrolling checklist

   Result saved to Drive content folder and linked in Publishing Tracker.
   ============================================================================ */

function formatThumbnailBrief() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.MASTER, "Stage 1 — Master Content")) return;

  const master = getMasterContent(idea.id);
  const ui     = SpreadsheetApp.getUi();

  if (!master || !master.thumbnailBrief) {
    ui.alert(
      "No Thumbnail Brief Found",
      "THUMBNAIL_BRIEF is empty for: " + idea.id + "\n\n" +
      "Run Stage 1 first — it generates the brief automatically.",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    "Stage 8E — Formatting Thumbnail Brief",
    "Claude is formatting the thumbnail brief for: " + idea.company +
    "\n\nOutputs:\n" +
    "• Canva spec (colors, fonts, layout)\n" +
    "• AI image generation prompt\n" +
    "• Stop-scrolling checklist\n\n" +
    "This takes ~15 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildThumbnailPrompt(idea, master);
    const raw    = callClaude(prompt, "stage_8_scenes");

    const docUrl = saveThumbnailBrief(idea, master, raw);
    writeThumbnailLinkToPublishing(idea.id, docUrl);

    ui.alert(
      "✅ Stage 8E Complete — Thumbnail Brief Ready",
      "Thumbnail brief saved to Drive and linked in Publishing Tracker.\n\n" +
      docUrl,
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 8E — Thumbnail", idea.id, "API Error", err.message);
    ui.alert("❌ Stage 8E Failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


function buildThumbnailPrompt(idea, master) {
  return `
You are formatting a YouTube thumbnail brief for GovernX — a business analytics
channel targeting C-Suite executives and business professionals.

CONTENT BRIEF:
Title          : ${master.title}
Company/Topic  : ${idea.company}
Discipline     : ${master.discipline}
Core Insight   : ${master.coreInsight}
Target Audience: ${master.targetAudience || "Business leaders, C-Suite"}

THUMBNAIL_BRIEF FROM STAGE 1:
${master.thumbnailBrief}

GOVERNX VISUAL IDENTITY (enforce strictly):
Background    : #0A0A0A (near-black)
Primary accent: #FF0000 (vivid red)
Text colors   : #FFFFFF (primary) / #CCCCCC (secondary)
Fonts         : Montserrat Bold / Black (English) | Cairo Bold (Arabic)
Style         : High contrast, no gradients, no soft colors, no decorative clutter
Dimensions    : 1280 × 720px (16:9 YouTube standard)

YOUR TASK — expand the THUMBNAIL_BRIEF above into three outputs:

CANVA_SPEC_START
TEXT_OVERLAY: [max 5 words — #FFFFFF or #FF0000]
BACKGROUND_TREATMENT: [background image + color overlay technique]
LAYOUT: [position of text / image / accent elements]
COLOR_PALETTE: [hex codes with purpose]
FONTS: [specific weight and size — e.g. Montserrat Black 96pt for headline]
VISUAL_ELEMENTS: [every element: image, icon, line, overlay]
DIMENSIONS: 1280 × 720px
CANVA_SPEC_END

AI_PROMPT_START
[Single paragraph for Midjourney/DALL-E/Ideogram.
 Include: subject, style, lighting, mood, color palette, composition.
 End with: --ar 16:9 --style raw
 NO TEXT in the image — text is added in Canva separately.]
AI_PROMPT_END

CHECKLIST_START
[ ] Text overlay is 5 words or fewer
[ ] Image communicates the story without text
[ ] Red (#FF0000) is present as a structural element
[ ] Viewer reads the thumbnail in under 0.5 seconds
[ ] Visual creates a question — not an answer
[ ] No generic stock photo feel — specific to this story
[ ] High contrast between text and background
[ ] GovernX brand recognizable at mobile thumbnail size
CHECKLIST_END

STOP_SCROLLING_ANALYSIS: [1–2 sentences — WHY will this thumbnail make someone stop?]
`;
}


function saveThumbnailBrief(idea, master, raw) {
  const getBlock = (s, e) => { const m = raw.match(new RegExp(s + "([\\s\\S]*?)" + e)); return m ? m[1].trim() : ""; };
  const getValue = (f)    => { const m = raw.match(new RegExp(f + ":\\s*([^\\n\\r]*)")); return m ? m[1].trim() : ""; };

  const canvaSpec = getBlock("CANVA_SPEC_START", "CANVA_SPEC_END");
  const aiPrompt  = getBlock("AI_PROMPT_START",  "AI_PROMPT_END");
  const checklist = getBlock("CHECKLIST_START",  "CHECKLIST_END");
  const analysis  = getValue("STOP_SCROLLING_ANALYSIS");
  const line = "═".repeat(60);

  let doc = line + "\nGOVERNX — THUMBNAIL BRIEF\n" + line + "\n";
  doc += "Content ID : " + idea.id + "\n";
  doc += "Title      : " + master.title + "\n";
  doc += "Generated  : " + new Date().toLocaleString() + "\n" + line + "\n\n";
  doc += "ORIGINAL BRIEF (from Stage 1):\n" + master.thumbnailBrief + "\n\n";
  doc += line + "\nSECTION 1 — CANVA SPEC\n" + line + "\n" + canvaSpec + "\n\n";
  doc += line + "\nSECTION 2 — AI IMAGE GENERATION PROMPT\n" + line + "\n" + aiPrompt + "\n\n";
  doc += line + "\nSECTION 3 — STOP-SCROLLING CHECKLIST\n" + line + "\n" + checklist + "\n\n";
  doc += line + "\nSTOP-SCROLLING ANALYSIS:\n" + analysis + "\n" + line + "\n";

  const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
  const docTitle      = idea.id + " — Thumbnail Brief";

  const existing = contentFolder.getFilesByName(docTitle);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const newDoc = DocumentApp.create(docTitle);
  newDoc.getBody().setText(doc);
  newDoc.saveAndClose();

  const file = DriveApp.getFileById(newDoc.getId());
  contentFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log("Thumbnail brief saved: " + newDoc.getUrl());
  return newDoc.getUrl();
}


function writeThumbnailLinkToPublishing(contentId, url) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) return;

  const pubData = pubSheet.getDataRange().getValues();
  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== contentId) continue;
    const existing = pubSheet.getRange(i + 1, COL_PUBLISHING.NOTES).getValue().toString();
    const tag      = "🖼️ Thumbnail Brief: " + url;
    if (!existing.includes("Thumbnail Brief:")) {
      pubSheet.getRange(i + 1, COL_PUBLISHING.NOTES)
        .setValue(existing ? existing + "\n" + tag : tag)
        .setWrap(true);
    }
    break;
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// STAGE 8E EXTENSION — Generate Thumbnail Image via Ideogram
//
// Reads the AI prompt from the Stage 8E brief doc (or re-derives it), calls
// the Ideogram v2 API to generate a 16:9 thumbnail image, saves the PNG to
// the Drive content folder, and uploads it to YouTube as the video thumbnail.
//
// Prerequisites:
//   • Script Property  IDEOGRAM_API_KEY  — from ideogram.ai
//   • Stage 8E complete (brief doc in Drive content folder)
//   • Stage 11 complete (YouTube Video ID in Publishing Tracker) — for upload
//
// Run from: GovernX menu → 🖼️ Generate Thumbnail Image (Ideogram)
// ══════════════════════════════════════════════════════════════════════════════
function generateThumbnailImage() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui      = SpreadsheetApp.getUi();
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const apiKey  = PropertiesService.getScriptProperties().getProperty("IDEOGRAM_API_KEY");

  if (!apiKey) {
    ui.alert("⚠️ Ideogram API Key Missing",
      "Add your key to Script Properties:\n" +
      "Key  : IDEOGRAM_API_KEY\nValue: your_key_from_ideogram.ai\n\n" +
      "Extensions → Apps Script → Project Settings → Script Properties",
      ui.ButtonSet.OK);
    return;
  }

  // ── Step 1: Get AI image prompt ───────────────────────────────────────────
  // Try to read from Stage 8E brief doc. Fall back to re-building from master.
  const master = getMasterContent(idea.id);
  if (!master) {
    ui.alert("⚠️ Master Content Not Found",
      "Run Stage 1 first to generate the Master Content entry for: " + idea.id,
      ui.ButtonSet.OK);
    return;
  }

  let imagePrompt = extractIdeogramPrompt_(idea.id, idea.company);

  if (!imagePrompt) {
    // Fall back: build prompt inline from thumbnailBrief if brief doc not found
    if (!master.thumbnailBrief) {
      ui.alert("⚠️ No Thumbnail Brief",
        "Run Stage 8E (Format Thumbnail Brief) first to generate the AI image prompt.",
        ui.ButtonSet.OK);
      return;
    }
    imagePrompt = buildFallbackIdeogramPrompt_(idea, master);
    Logger.log("Stage 8E/Ideogram: Using fallback prompt from thumbnailBrief.");
  }

  // ── Step 2: Confirm with user ─────────────────────────────────────────────
  const confirm = ui.alert(
    "🖼️ Generate Thumbnail — Ideogram",
    "Ready to generate thumbnail for: " + idea.company + "\n\n" +
    "Prompt preview (first 200 chars):\n" + imagePrompt.substring(0, 200) + "…\n\n" +
    "Model   : Ideogram V2\n" +
    "Format  : 16:9  (1280 × 720)\n" +
    "Style   : Realistic\n\n" +
    "The image will be saved to Drive and optionally uploaded to YouTube.\n\n" +
    "Proceed?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    // ── Step 3: Call Ideogram API ─────────────────────────────────────────
    Logger.log("Stage 8E/Ideogram: Calling API for " + idea.id);

    const ideogramPayload = {
      image_request: {
        prompt      : imagePrompt,
        aspect_ratio: "ASPECT_16_9",
        model       : "V_2",
        style_type  : "REALISTIC",
        num_images  : 1
      }
    };

    const ideogramResponse = UrlFetchApp.fetch("https://api.ideogram.ai/generate", {
      method            : "post",
      contentType       : "application/json",
      headers           : { "Api-Key": apiKey },
      payload           : JSON.stringify(ideogramPayload),
      muteHttpExceptions: true
    });

    const ideogramCode = ideogramResponse.getResponseCode();
    const ideogramBody = JSON.parse(ideogramResponse.getContentText());

    if (ideogramCode !== 200) {
      throw new Error("Ideogram API error " + ideogramCode + ": " +
        JSON.stringify(ideogramBody).substring(0, 300));
    }

    const imageUrl = ideogramBody.data && ideogramBody.data[0]
      ? ideogramBody.data[0].url
      : null;

    if (!imageUrl) {
      throw new Error("Ideogram returned no image URL. Response: " +
        JSON.stringify(ideogramBody).substring(0, 300));
    }

    Logger.log("Stage 8E/Ideogram: Image generated → " + imageUrl);

    // ── Step 4: Download and save to Drive ───────────────────────────────
    const imgResponse = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    if (imgResponse.getResponseCode() !== 200) {
      throw new Error("Could not download generated image from Ideogram CDN.");
    }

    const imgBlob       = imgResponse.getBlob().setContentType("image/png");
    const imgFileName   = idea.id + " — Thumbnail.png";
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);

    // Replace any existing thumbnail PNG
    const existingImgs = contentFolder.getFilesByName(imgFileName);
    while (existingImgs.hasNext()) existingImgs.next().setTrashed(true);

    const imgFile    = contentFolder.createFile(imgBlob.setName(imgFileName));
    const imgDriveUrl = "https://drive.google.com/file/d/" + imgFile.getId() + "/view";
    Logger.log("Stage 8E/Ideogram: Saved to Drive → " + imgDriveUrl);

    // ── Step 5: Write Drive link to Publishing Tracker (THUMBNAIL col) ───
    writeThumbnailImageLinkToPublishing_(idea.id, imgDriveUrl);

    // ── Step 6: Optionally upload to YouTube as thumbnail ─────────────────
    const videoId = getYouTubeVideoId_(ss, idea.id);
    let ytUploaded = false;

    if (videoId) {
      const ytConfirm = ui.alert(
        "📺 Upload to YouTube as Thumbnail?",
        "Video ID found: " + videoId + "\n\n" +
        "Upload the generated image as the YouTube thumbnail now?\n\n" +
        "(You can also set it manually in YouTube Studio.)",
        ui.ButtonSet.YES_NO
      );

      if (ytConfirm === ui.Button.YES) {
        try {
          // Re-fetch as JPEG — YouTube requires JPEG or PNG ≤ 2MB
          const ytBlob = imgFile.getBlob().setContentType("image/png");
          YouTube.Thumbnails.set(videoId, ytBlob);
          ytUploaded = true;
          Logger.log("Stage 8E/Ideogram: Thumbnail uploaded to YouTube for " + videoId);
        } catch (ytErr) {
          Logger.log("Stage 8E/Ideogram: YouTube thumbnail upload failed (non-fatal): " + ytErr.message);
          ui.alert("⚠️ YouTube Upload Failed",
            "The image was saved to Drive but could not be set as YouTube thumbnail:\n\n" +
            ytErr.message + "\n\nSet it manually in YouTube Studio.",
            ui.ButtonSet.OK);
        }
      }
    }

    ui.alert(
      "✅ Thumbnail Image Generated",
      "AI thumbnail created for: " + idea.company + "\n\n" +
      "📁 Saved to Drive: " + imgDriveUrl + "\n\n" +
      (ytUploaded
        ? "📺 Uploaded to YouTube as thumbnail ✅"
        : videoId
          ? "📺 Not uploaded to YouTube — set manually in YouTube Studio."
          : "📺 No YouTube Video ID yet — run Stage 11 first, then re-run to upload.") + "\n\n" +
      "To regenerate, run this function again — existing file will be replaced.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 8E — Thumbnail Image", idea.id, "Ideogram/API Error", err.message);
    Logger.log("Stage 8E/Ideogram ERROR: " + err.message);
    ui.alert("❌ Thumbnail Generation Failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


// ── Reads the AI_PROMPT block from the Stage 8E brief Google Doc ─────────────
function extractIdeogramPrompt_(contentId, company) {
  try {
    const rootSearch = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    if (!rootSearch.hasNext()) return null;
    const rootFolder = rootSearch.next();

    const subName   = contentId + " — " + company;
    const subSearch = rootFolder.getFoldersByName(subName);
    if (!subSearch.hasNext()) return null;
    const contentFolder = subSearch.next();

    const docTitle  = contentId + " — Thumbnail Brief";
    const docSearch = contentFolder.getFilesByName(docTitle);
    if (!docSearch.hasNext()) return null;
    const docFile   = docSearch.next();

    const doc  = DocumentApp.openById(docFile.getId());
    const text = doc.getBody().getText();

    // Extract between AI_PROMPT_START and AI_PROMPT_END markers
    const match = text.match(/AI_PROMPT_START\s*([\s\S]*?)\s*AI_PROMPT_END/);
    if (!match || !match[1].trim()) return null;

    return match[1].trim();
  } catch (e) {
    Logger.log("extractIdeogramPrompt_: " + e.message);
    return null;
  }
}


// ── Fallback prompt builder when Stage 8E brief doc is not available ──────────
function buildFallbackIdeogramPrompt_(idea, master) {
  const brief = master.thumbnailBrief || "";
  // Parse TEXT, IMAGE, COLOR, HOOK fields from thumbnailBrief if present
  const getField = (f) => {
    const m = brief.match(new RegExp(f + ":\\s*([^|\\n]+)"));
    return m ? m[1].trim() : "";
  };
  const text  = getField("TEXT")  || master.coreInsight || "";
  const image = getField("IMAGE") || "corporate executive boardroom, dramatic lighting";
  const color = getField("COLOR") || "dark background, vivid red accent";

  return (
    "YouTube thumbnail for a business analytics video about " + idea.company + ". " +
    (text  ? "The story: " + text + ". " : "") +
    (image ? "Visual concept: " + image + ". " : "") +
    "Style: " + color + ", cinematic, high contrast, photorealistic. " +
    "No text overlay — this is the background image only. " +
    "GovernX brand colors: #0A0A0A near-black background, #FF0000 red accent. " +
    "Composition: wide shot, dramatic perspective. --ar 16:9 --style raw"
  );
}


// ── Write thumbnail image Drive link to Publishing Tracker col 4 (THUMBNAIL) ──
function writeThumbnailImageLinkToPublishing_(contentId, driveUrl) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) return;

  const pubData = pubSheet.getDataRange().getValues();
  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== contentId) continue;
    pubSheet.getRange(i + 1, COL_PUBLISHING.THUMBNAIL)
      .setFormula('=HYPERLINK("' + driveUrl + '","🖼️ Thumbnail")');
    SpreadsheetApp.flush();
    return;
  }
}


// ── Get YouTube Video ID from Publishing Tracker ──────────────────────────────
function getYouTubeVideoId_(ss, contentId) {
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) return null;
  const pubData = pubSheet.getDataRange().getValues();
  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== contentId) continue;
    const vid = pubData[i][COL_PUBLISHING.YOUTUBE_VIDEO_ID - 1].toString().trim();
    return vid.length > 5 ? vid : null;
  }
  return null;
} 