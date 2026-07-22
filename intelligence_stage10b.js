/* ============================================================================
   Intelligence_4.2_Stage10B.gs — GovernX Intelligence Platform
   PHASE 4 · UNIT 4.2 — Stage 10B: Commercial Layer

   For a given video, Claude recommends genuinely relevant monetization assets
   (books, courses, software, templates) and drafts a Newsletter CTA + a Premium
   Report CTA, writing it all into Affiliate_Assets (Unit 4.1).

   PIPELINE POSITION: runs after Stage 10 (metadata) / Stage 11 (upload).
     YouTube Upload → [Stage 10B] → Newsletter · Affiliate · Reports · Courses

   SAFETY: additive, own constants, intelSS_(), no config.gs edits.
   Reuses: getSelectedVideoId_ (Unit 3.2), COL_AFFILIATE (Unit 4.1), and reads
   the production Master Content Table (SHEET.MASTER / COL_MASTER from config.gs).

   HOW TO USE:
   1. Ensure setupRevenueTabs() has been run.
   2. Select a video row (ID in column A), run  generateCommercialLayer().
   ============================================================================ */


const COMMERCIAL_SYSTEM_CONTEXT = `
You are a monetization strategist for GovernX, a governance/business analysis
channel. You recommend REAL, well-known, genuinely relevant resources tied to a
video's governance/business lesson. Never invent fake products or courses. Keep
recommendations credible and specific. Draft crisp, non-hypey calls to action.
`;


// ══════════════════════════════════════════════════════════════════════════════
// STAGE 10B — generate the commercial layer for one video
// ══════════════════════════════════════════════════════════════════════════════
function generateCommercialLayer(videoId) {
  const prodSS = SpreadsheetApp.getActiveSpreadsheet();  // production tabs
  const ss     = intelSS_();                              // intelligence tabs
  const ui     = SpreadsheetApp.getUi();

  if (!videoId) videoId = getSelectedVideoId_();          // from Unit 3.2
  if (!videoId) return;

  const aff = ss.getSheetByName(SHEET_REVENUE.AFFILIATE);
  if (!aff) { ui.alert("Run setupRevenueTabs() first (Affiliate_Assets missing)."); return; }

  // ── Pull the video's topic from Master Content Table ──────────────────────
  let title = "", discipline = "", field = "", industry = "", insight = "";
  const master = prodSS.getSheetByName(SHEET.MASTER);
  if (master) {
    const data = master.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][COL_MASTER.ID - 1] || "").toString().trim() === videoId) {
        title      = data[i][COL_MASTER.TITLE        - 1];
        discipline = data[i][COL_MASTER.DISCIPLINE   - 1];
        field      = data[i][COL_MASTER.FIELD        - 1];
        industry   = data[i][COL_MASTER.INDUSTRY     - 1];
        insight    = data[i][COL_MASTER.CORE_INSIGHT - 1];
        break;
      }
    }
  }
  if (!title) title = videoId;   // fallback

  const prompt = `
Recommend the commercial layer for this GovernX video.

Video   : ${title}
Discipline: ${discipline || "GRC/BPR"}
Industry: ${industry || ""} ${field ? "(" + field + ")" : ""}
Lesson  : ${insight || ""}

Recommend REAL, relevant resources a viewer of this video would value. Return
EXACTLY these fields, one per line (use "; " to separate multiple items):

BOOKS: [2-3 real books — "Title — Author"]
COURSES: [1-2 relevant courses/platforms]
SOFTWARE: [1-2 relevant tools, e.g. GRC/audit/process software]
TEMPLATES: [1-2 practical templates a viewer could use]
NEWSLETTER_CTA: [1 line inviting viewers to the GovernX newsletter]
REPORT_CTA: [1 line offering a premium governance report on this topic]
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, COMMERCIAL_SYSTEM_CONTEXT, "medium", 1500);
    const g = (f) => { const m = raw.match(new RegExp(f + ":\\s*(.+)")); return m ? m[1].trim() : ""; };

    const books     = g("BOOKS");
    const courses   = g("COURSES");
    const software  = g("SOFTWARE");
    const templates = g("TEMPLATES");
    const cta       = [g("NEWSLETTER_CTA"), g("REPORT_CTA")].filter(Boolean).join("  |  ");

    // Upsert the Affiliate_Assets row for this video
    let row = findAffiliateRow_(aff, videoId);
    if (row === -1) row = aff.getLastRow() + 1;

    aff.getRange(row, COL_AFFILIATE.CONTENT_ID    ).setValue(videoId);
    aff.getRange(row, COL_AFFILIATE.BOOKS         ).setValue(books);
    aff.getRange(row, COL_AFFILIATE.COURSES       ).setValue(courses);
    aff.getRange(row, COL_AFFILIATE.SOFTWARE      ).setValue(software);
    aff.getRange(row, COL_AFFILIATE.TEMPLATES     ).setValue(templates);
    aff.getRange(row, COL_AFFILIATE.CTA           ).setValue(cta);
    // Affiliate_Link + Revenue left for you to fill (placeholder if empty)
    if (!aff.getRange(row, COL_AFFILIATE.AFFILIATE_LINK).getValue()) {
      aff.getRange(row, COL_AFFILIATE.AFFILIATE_LINK).setValue("[paste your affiliate links]");
    }

    ui.alert("✅ Commercial layer generated for " + videoId,
      "Books: " + books + "\nCourses: " + courses + "\nSoftware: " + software +
      "\nTemplates: " + templates + "\n\nCTAs written. Add your real affiliate links in the Affiliate_Link column.",
      ui.ButtonSet.OK);

  } catch (err) {
    if (typeof logError === "function") logError("Stage 10B — Commercial", videoId, "API/Runtime", err.message);
    ui.alert("❌ Stage 10B failed: " + err.message);
  }
}

function findAffiliateRow_(aff, videoId) {
  const data = aff.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_AFFILIATE.CONTENT_ID - 1] || "").toString().trim() === videoId) return i + 1;
  }
  return -1;
}
