/* ============================================================================
   Stage_8E_Thumbnail_Remotion.gs — GovernX Content OS

   Renders the YouTube thumbnail in Remotion instead of generating it with an
   image model, and sets it on the uploaded video.

   Why the change:
     • TEXT. Diffusion models still garble headlines — a generated reference
       poster came back reading "SCCAN DAL". A thumbnail with a typo is a brand
       problem you have to regenerate around.
     • BRAND. Every generation drifts off the design system. Rendering with the
       same components and tokens as the film means the thumbnail always matches
       the video it sits next to.
     • TRUTH. The headline figure is pulled from the same verified data moment
       the film uses, so the number on the thumbnail cannot disagree with the
       number in the video.
     • It is free and takes about a minute — no image-model key.

   The one thing this cannot invent is a photograph. Put a real image (or an
   AI-generated background with NO TEXT in it) in the Idea Catalogue Note as
   `thumb_photo=<url>` and it is composited under the typography.
   ============================================================================ */

// Publishing Tracker column that holds the thumbnail link (COL_PUBLISHING.THUMBNAIL = 4)
function generateThumbnailRemotion() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = getMasterContent(idea.id) || {};

  // ── Headline figure: the strongest verified data moment ───────────────────
  const hero = thumbHeroFigure_(idea.id);
  if (!hero.value) {
    ui.alert("No figure to put on the thumbnail",
      "No data moments found for " + idea.id + ".\n\n" +
      "Run the research steps (② → ③) first — the thumbnail uses the same\n" +
      "verified figure as the film so the two can never disagree.",
      ui.ButtonSet.OK);
    return;
  }

  // The thumbnail is deterministic (same verified data → same poster). To get a
  // DIFFERENT one, edit the style/copy in this prompt — don't just re-run. Three
  // styles: dark poster, paper "leaked document", cinematic (courthouse depth).
  const dHeadline = thumbHeadline_(master.title || idea.company);
  const dBanner   = String(master.discipline ? master.discipline + " · GOVERNANCE FAILURE" : "GOVERNANCE FAILURE");
  const dDate     = String(master.checkpointDate || hero.dateTag || "").toUpperCase();

  const resp = ui.prompt("🖼️ Thumbnail — style & copy — " + idea.id,
    "Press OK for the defaults, or edit any field to get a different thumbnail.\n\n" +
    "Format:  style | headline | value | unit | banner\n" +
    "style = dark | paper | cinematic\n" +
    "(cinematic: 'headline' becomes the 2nd big line, keep it short; 'banner' is the gold label)\n\n" +
    "dark | " + dHeadline + " | " + hero.value + " | " + hero.unit + " | " + dBanner,
    ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const raw   = resp.getResponseText().trim();
  const parts = (raw || ("dark|" + dHeadline + "|" + hero.value + "|" + hero.unit + "|" + dBanner)).split("|");
  const style    = (parts[0] || "dark").trim().toLowerCase();
  const headline = (parts[1] || dHeadline).trim();
  const value    = (parts[2] || hero.value).trim();
  const unit     = (parts[3] || hero.unit).trim();
  const banner   = (parts[4] || dBanner).trim();

  let compositionId, props;
  if (style.indexOf("cinema") === 0) {
    compositionId = "ThumbnailCinematic";
    props = {
      masthead: "GOVERNANCE AUDIT · " + String(idea.company || "").toUpperCase(),
      caseTag : "CASE FILE Nº " + idea.id,
      line1   : "THE " + value,
      line2   : headline.toUpperCase(),
      subline : banner.toUpperCase()
    };
  } else {
    compositionId = "ThumbnailPoster";
    props = {
      company   : String(idea.company || "").toUpperCase(),
      headline  : headline,
      bigValue  : value,
      bigUnit   : unit,
      hedge     : hero.hedge,
      caption   : hero.caption,
      dateTag   : dDate,
      bannerText: banner,
      photoSrc  : thumbPhotoFromNote_(idea),
      ground    : (style.indexOf("paper") === 0 ? "paper" : "navy")
    };
  }
  const variant = compositionId === "ThumbnailCinematic" ? "cinematic"
                : (props.ground === "paper" ? "paper" : "dark");

  try {
    const serverBase = getRemotionServerUrl();
    const res = UrlFetchApp.fetch(serverBase + "/thumbnail", {
      method            : "post",
      contentType       : "application/json",
      headers           : { "ngrok-skip-browser-warning": "true" },
      payload           : JSON.stringify({ contentId: idea.id, compositionId: compositionId, variant: variant, props: props }),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    let body;
    try { body = JSON.parse(res.getContentText()); }
    catch (e) {
      throw new Error("Server did not return JSON (HTTP " + code + "). " +
        "Is the Remotion server running, and is REMOTION_SERVER_URL current? " +
        "Use 🩺 Check render engine is reachable.");
    }
    if (code !== 200 || !body.success) throw new Error(body.error || "Render failed (HTTP " + code + ")");

    // Download the PNG through the same base the server is reachable on.
    const pngUrl = String(body.url)
      .replace(/https?:\/\/localhost:\d+/, serverBase)
      .replace(/https?:\/\/127\.0\.0\.1:\d+/, serverBase);
    const png = UrlFetchApp.fetch(pngUrl, {
      muteHttpExceptions: true, headers: { "ngrok-skip-browser-warning": "true" }
    });
    // Gate on real bytes, NOT the Content-Type header. UrlFetchApp's getHeaders()
    // is case-inconsistent ("Content-Type" vs "content-type"), so a good 200 + PNG
    // was being rejected as "HTTP 200, " with an empty type. The blob is forced to
    // image/png on save below, so the header is irrelevant.
    const pngBytes = png.getResponseCode() === 200 ? png.getBlob().getBytes().length : 0;
    if (pngBytes < 200) {
      throw new Error("Could not download the PNG (HTTP " + png.getResponseCode() +
        ", " + pngBytes + " bytes). Is the Remotion server still running?");
    }

    // Keep the three styles side by side in a Thumbnails subfolder so you can
    // compare before choosing. Re-rendering the SAME style replaces only that one
    // (filename carries the variant); the other styles are left untouched.
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
    const thumbsIt = contentFolder.getFoldersByName("Thumbnails");
    const thumbsFolder = thumbsIt.hasNext() ? thumbsIt.next() : contentFolder.createFolder("Thumbnails");
    const existing = thumbsFolder.getFilesByName(body.filename);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const file = thumbsFolder.createFile(png.getBlob().setName(body.filename).setContentType("image/png"));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = file.getUrl();

    writePublishingLink(idea.id, idea.company, master,
      COL_PUBLISHING.THUMBNAIL, url, "🖼️ Thumbnail");

    SpreadsheetApp.flush();
    ui.alert("✅ Thumbnail rendered — " + variant + " style",
      body.filename + "\n\n" + url + "\n\n" +
      "Saved in the content folder → Thumbnails subfolder (the three styles sit\n" +
      "side by side there). Latest render is linked in the Publishing Tracker.\n" +
      "Re-run and pick dark / paper / cinematic to add the other styles.",
      ui.ButtonSet.OK);

  } catch (err) {
    logError("Stage 8E — Thumbnail (Remotion)", idea.id, "Render Error", err.message);
    ui.alert("❌ Thumbnail failed", err.message + "\nSee Error Log.", ui.ButtonSet.OK);
  }
}


/* ── helpers ────────────────────────────────────────────────────────────────── */

// The strongest verified figure, taken from Research_Data_Moments so the
// thumbnail and the film quote the same number. Prefers a short, punchy value.
function thumbHeroFigure_(contentId) {
  const out = { value: "", unit: "", hedge: "", caption: "", source: "", dateTag: "" };
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RB.DM);
    if (!sh || sh.getLastRow() < 2) return out;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, COL_RB_DM.length).getValues();

    let best = null, bestScore = -1;
    rows.forEach(function (r) {
      const raw = String(r[2] || "").trim();          // Value
      if (!raw || !/\d/.test(raw)) return;            // a thumbnail needs a number
      let score = 0;
      if (/million|billion|bn\b|[$€£¥]/i.test(raw)) score += 3;   // money reads big
      if (raw.length <= 18) score += 2;                          // fits the poster
      if (/regulat|court/i.test(String(r[5] || ""))) score += 2;  // official
      if (score > bestScore) { bestScore = score; best = r; }
    });
    if (!best) return out;

    const raw = String(best[2]).trim();
    // "more than $140 million" → hedge "MORE THAN", value "$140", unit "MILLION"
    const hedge = (raw.match(/^(more than|approximately|roughly|at least|over|nearly|about)\b/i) || [])[0] || "";
    const rest  = raw.replace(/^(more than|approximately|roughly|at least|over|nearly|about)\s*/i, "").trim();
    const m     = rest.match(/^([$€£¥]?\s?[\d.,]+)\s*(.*)$/);

    out.hedge   = hedge.toUpperCase();
    out.value   = m ? m[1].replace(/\s/g, "") : rest;
    out.unit    = m ? String(m[2] || "").toUpperCase() : "";
    out.caption = String(best[3] || "").toUpperCase().slice(0, 58);
    out.source  = String(best[4] || "").replace(/^Source:\s*/i, "");
  } catch (e) { /* fall through with blanks */ }
  return out;
}

// Title → an all-caps poster headline, trimmed to something that can be set big.
function thumbHeadline_(title) {
  let t = String(title || "").trim();
  t = t.split(/[:—–]/)[0].trim();                    // drop the subtitle
  t = t.replace(/^how\s+/i, "").trim();              // "How X did Y" → "X did Y"
  if (t.length > 44) {
    const cut = t.slice(0, 44);
    t = cut.slice(0, Math.max(cut.lastIndexOf(" "), 26));
  }
  return t.toUpperCase();
}

// The thumbnail link is written with writePublishingLink, which stores a
// HYPERLINK formula — so the plain cell value is the label, not the URL. Read the
// formula when the value is not itself a link. Used by Stage 11.
function readPublishingThumbnailUrl_(pubSheet, pubRow) {
  if (!pubSheet || !pubRow || pubRow < 2) return "";
  const cell = pubSheet.getRange(pubRow, COL_PUBLISHING.THUMBNAIL);
  const raw = String(cell.getValue() || "").trim();
  if (/^https?:\/\//i.test(raw)) return raw;
  const m = String(cell.getFormula() || "").match(/HYPERLINK\("(https?:[^"]+)"/i);
  return m ? m[1] : "";
}

// Optional real image: put `thumb_photo=<url>` anywhere in the Idea Catalogue Note.
// Diffusion models are good at imagery and bad at text — generate a background
// with NO words in it, drop the link here, and let Remotion set the type.
function thumbPhotoFromNote_(idea) {
  try {
    const note = String(idea.note || "");
    const m = note.match(/thumb_photo\s*=\s*(\S+)/i);
    return m ? m[1] : "";
  } catch (e) { return ""; }
}
