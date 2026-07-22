/* ============================================================================
   Intelligence_4.1_Revenue.gs — GovernX Intelligence Platform
   PHASE 4 · UNIT 4.1 — Revenue Engine tables

   Turns the pipeline from "ends at YouTube upload" into a revenue funnel:
     YouTube → Newsletter → Affiliate → Reports → Courses → Consulting

   Tabs created:
     • Affiliate_Assets  (per-video monetization: books/courses/software/…, CTA, revenue)
     • Digital_Products  (your own products: reports/courses/templates/consulting)

   Unit 4.1 = the tables + setup. Unit 4.2 (Stage 10B) auto-fills Affiliate_Assets
   per video from its topic/company.

   SAFETY: additive, own constants, intelSS_(), no config.gs edits.
   Tabs pre-registered in INTEL_TABS → covered by organize/hide/show.

   HOW TO USE:
   1. Run  setupRevenueTabs()   → creates both tabs.
   2. Run  addDigitalProduct()  → add one of your own products (interactive).
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_REVENUE = {
  AFFILIATE : "Affiliate_Assets",
  PRODUCTS  : "Digital_Products"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_AFFILIATE = {
  CONTENT_ID     : 1,   // video ID (GX-…)
  BOOKS          : 2,
  COURSES        : 3,
  SOFTWARE       : 4,
  TEMPLATES      : 5,
  AFFILIATE_LINK : 6,
  CTA            : 7,
  REVENUE        : 8
};

const COL_PRODUCTS = {
  PRODUCT_ID    : 1,   // PR-###
  TYPE          : 2,
  PRICE         : 3,
  LINKED_VIDEOS : 4
};

const PRODUCT_TYPE_ENUM = ["Report", "Course", "Template", "Newsletter", "Consulting"];

const INTEL_HEADER_BG_RV = "#1a1a2e";
const INTEL_HEADER_FG_RV = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupRevenueTabs() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  // ── Affiliate_Assets ──────────────────────────────────────────────────────
  let aff = ss.getSheetByName(SHEET_REVENUE.AFFILIATE);
  const affNew = !aff;
  if (affNew) aff = ss.insertSheet(SHEET_REVENUE.AFFILIATE);
  aff.getRange(1, 1, 1, 8).setValues([[
    "Content_ID", "Books", "Courses", "Software", "Templates",
    "Affiliate_Link", "CTA", "Revenue"
  ]]).setBackground(INTEL_HEADER_BG_RV).setFontColor(INTEL_HEADER_FG_RV).setFontWeight("bold");
  aff.setFrozenRows(1);
  [160, 240, 240, 220, 220, 260, 300, 110].forEach((w, i) => aff.setColumnWidth(i + 1, w));

  // ── Digital_Products ──────────────────────────────────────────────────────
  let prod = ss.getSheetByName(SHEET_REVENUE.PRODUCTS);
  const prodNew = !prod;
  if (prodNew) prod = ss.insertSheet(SHEET_REVENUE.PRODUCTS);
  prod.getRange(1, 1, 1, 4).setValues([["Product_ID", "Type", "Price", "Linked_Videos"]])
      .setBackground(INTEL_HEADER_BG_RV).setFontColor(INTEL_HEADER_FG_RV).setFontWeight("bold");
  prod.setFrozenRows(1);
  [120, 140, 100, 360].forEach((w, i) => prod.setColumnWidth(i + 1, w));
  prod.getRange(2, COL_PRODUCTS.TYPE, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PRODUCT_TYPE_ENUM, true).setAllowInvalid(true).build());

  ui.alert("✅ Revenue Tabs Ready",
    (affNew ? "Created Affiliate_Assets. " : "Affiliate_Assets refreshed. ") +
    (prodNew ? "Created Digital_Products.\n\n" : "Digital_Products refreshed.\n\n") +
    "Affiliate_Assets auto-fills per video in Stage 10B (Unit 4.2).\n" +
    "Add your own products with addDigitalProduct().",
    ui.ButtonSet.OK);
}


// ── Next PR-### id ────────────────────────────────────────────────────────────
function generateProductId_(prodSheet) {
  const data = prodSheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const m = (data[i][COL_PRODUCTS.PRODUCT_ID - 1] || "").toString().match(/^PR-(\d+)$/);
    if (m) { const s = parseInt(m[1], 10); if (s > maxSeq) maxSeq = s; }
  }
  return "PR-" + String(maxSeq + 1).padStart(3, "0");
}


// ══════════════════════════════════════════════════════════════════════════════
// ADD A DIGITAL PRODUCT (interactive)
// ══════════════════════════════════════════════════════════════════════════════
function addDigitalProduct() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();
  const prod = ss.getSheetByName(SHEET_REVENUE.PRODUCTS);
  if (!prod) { ui.alert("Run setupRevenueTabs() first."); return; }

  const typeResp = ui.prompt("New Product — Type",
    "Type (" + PRODUCT_TYPE_ENUM.join(" / ") + "):", ui.ButtonSet.OK_CANCEL);
  if (typeResp.getSelectedButton() !== ui.Button.OK) return;
  const type = typeResp.getResponseText().trim() || "Report";

  const priceResp = ui.prompt("New Product — Price",
    "Price (number, e.g. 49):", ui.ButtonSet.OK_CANCEL);
  if (priceResp.getSelectedButton() !== ui.Button.OK) return;
  const price = priceResp.getResponseText().trim();

  const linkResp = ui.prompt("New Product — Linked Videos",
    "Linked video IDs (comma-separated, optional):", ui.ButtonSet.OK_CANCEL);
  if (linkResp.getSelectedButton() !== ui.Button.OK) return;
  const linked = linkResp.getResponseText().trim();

  const id = generateProductId_(prod);
  prod.appendRow([id, type, price, linked]);
  ui.alert("✅ Added " + id + " (" + type + ").");
}
