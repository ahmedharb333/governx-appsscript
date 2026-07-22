/* ============================================================================
   id_generator.gs — GovernX Content OS
   Auto-generates GX-YYMM-DOMAIN-SEQ ID when a row is filled in Idea Catalogue
   Format: GX-2605-TECH-001
   ============================================================================ */

// ── onEdit trigger — auto-generates the ID for any row that has a Company ─────
// Fires on an edit to ANY column, not just Company. The old version required the
// Company cell itself to be the edited one, so filling the columns in a different
// order — or pasting a block of rows — left rows with no ID and forced a manual
// fallback. Every row the edit touched is checked, so a multi-row paste is
// covered too.
//
// NOTE: a simple onEdit CANNOT see programmatic writes (setValue from a script),
// so a row created by Stage 0 still arrives without an ID. That case is handled
// by getActiveIdeaRow() in pipeline.gs, which generates on demand. Between the
// two, no row can reach a stage without an ID — which is why the manual
// "Generate ID" menu item was removed.
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET.IDEA) return;

    const first = e.range.getRow();
    const last  = first + e.range.getNumRows() - 1;

    for (let row = Math.max(first, 2); row <= last; row++) {
      const existingId = sheet.getRange(row, COL_IDEA.ID).getValue();
      if (existingId && existingId.toString().trim() !== "") continue;

      const company = sheet.getRange(row, COL_IDEA.COMPANY).getValue();
      if (!company || company.toString().trim() === "") continue;

      // Only mint an ID when the domain is actually KNOWN. A simple trigger has
      // no reliable UI to ask, so an unresolved domain is left alone rather than
      // silently stamped BIZ — getActiveIdeaRow() prompts for it on the first
      // stage run, and the ID is created then.
      const resolved = resolveDomainCode_(sheet, row);
      if (!resolved) {
        Logger.log("Row " + row + ": domain unknown for \"" + company + "\" — ID deferred to first stage run");
        continue;
      }

      const newId = generateId(sheet, row, resolved.code);
      sheet.getRange(row, COL_IDEA.ID).setValue(newId);
      Logger.log("Generated ID: " + newId + " for row " + row + " — domain from " + resolved.how);
    }

  } catch (err) {
    // An onEdit failure must never block the user's typing.
    try { logError("ID Generator", "", "Auto-trigger", err.message); } catch (e2) {}
  }
}

/* ── Domain resolution ────────────────────────────────────────────────────────
   The ID is the join key for every sheet, Drive folder and tracker row, so the
   domain code baked into it is expensive to change later. It must never be a
   silent guess.

   Two vocabularies exist and neither matches DOMAIN_CODES:
     • Idea Catalogue "Domain" holds things like "Private Equity", "Automotive" —
       not DOMAIN_CODES keys, so `DOMAIN_CODES[domain] || "BIZ"` quietly returned
       BIZ for both. They were right by luck, not by decision.
     • Company_Master "Domain" is really an INDUSTRY ("Technology", "Finance",
       "Crypto", "Sports Business" — 68 distinct values), so a direct lookup would
       map Technology → BIZ when it should be TECH.

   domainToCode_ normalises BOTH vocabularies by keyword. resolveDomainCode_
   returns null when it genuinely cannot tell — callers must then ask rather than
   invent a code. */
function domainToCode_(raw) {
  const s = String(raw || "").trim();
  if (!s) return null;
  if (DOMAIN_CODES[s]) return DOMAIN_CODES[s];        // exact GovernX domain
  const t = s.toLowerCase();

  // order matters: the most specific signal wins
  if (/\bsport|football|league|athlet|olympic/.test(t))                   return "SPT";
  if (/\bmedia|entertain|gaming|creator|publish|broadcast|film|music/.test(t)) return "MED";
  // PUB is the PUBLIC SECTOR — government bodies, ministries, municipalities,
  // state-owned entities. NOT private firms that merely serve or are regulated by
  // government: a defence contractor and a listed utility are businesses, and
  // classing them PUB would misfile the majority of Company_Master.
  if (/\bpublic sector|government|govt|ministry|municipal|state-owned|regulator|central bank/.test(t)) return "PUB";
  if (/\btech|software|saas|semiconductor|fintech|crypto|digital|data|cloud|ai\b|internet|mobility|imaging|health tech/.test(t)) return "TECH";
  // finance, retail, energy, industrial, healthcare, automotive, aerospace,
  // utilities, real estate… all sit under Business in the GovernX taxonomy
  if (/\bfinanc|bank|insur|invest|private equity|retail|consumer|energy|oil|gas|mining|industrial|manufactur|automotive|aerospace|defen[cs]e|utilit|health|pharma|food|beverage|real estate|construction|chemical|transport|aviation|airline|telecom|conglomerate|logistics|hospitality|travel|agri|professional services|natural resources|environmental/.test(t)) return "BIZ";
  return null;                                        // unknown — do NOT guess
}

// Look a company up in Company_Master and translate its industry to a code.
function domainCodeFromCompanyMaster_(companyName) {
  const name = String(companyName || "").trim().toLowerCase();
  if (!name) return null;
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Company_Master");
    if (!sh || sh.getLastRow() < 2) return null;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();  // A..E
    for (let i = 0; i < data.length; i++) {
      const cm = String(data[i][1] || "").trim().toLowerCase();          // B Company_Name
      if (!cm) continue;
      // tolerate "Kodac" vs "Kodak Company" style differences
      if (cm === name || cm.indexOf(name) === 0 || name.indexOf(cm) === 0) {
        const code = domainToCode_(data[i][4]);                          // E Domain
        if (code) return { code: code, industry: String(data[i][4]).trim(), matched: String(data[i][1]).trim() };
      }
    }
  } catch (e) { /* Company_Master absent — fall through */ }
  return null;
}

/** → { code, how } or null when the domain genuinely cannot be determined. */
function resolveDomainCode_(sheet, row) {
  const raw = sheet.getRange(row, COL_IDEA.DOMAIN).getValue();
  const fromCol = domainToCode_(raw);
  if (fromCol) return { code: fromCol, how: 'Domain column ("' + String(raw).trim() + '")' };

  const company = sheet.getRange(row, COL_IDEA.COMPANY).getValue();
  const fromCM  = domainCodeFromCompanyMaster_(company);
  if (fromCM) return { code: fromCM.code, how: 'Company_Master — ' + fromCM.matched + " (" + fromCM.industry + ")" };

  return null;
}

// ── Core ID generation logic ──────────────────────────────────────────────────
// `domCode` may be passed in by a caller that already resolved (and possibly
// asked the user for) the domain; otherwise resolve here. BIZ remains the last
// resort, but resolveDomainCode_ now covers both vocabularies so the fallback is
// reached far less often — and callers can detect null and ask first.
function generateId(sheet, row, domCodeIn) {

  const resolved = domCodeIn || (resolveDomainCode_(sheet, row) || {}).code || "BIZ";
  const domCode  = resolved;

  // Build YYMM from today's date
  const now   = new Date();
  const yy    = now.getFullYear().toString().slice(2);        // "26"
  const mm    = String(now.getMonth() + 1).padStart(2, "0"); // "05"
  const datePart = yy + mm;                                  // "2605"

  // Build prefix to match against existing IDs for sequence
  const prefix = "GX-" + datePart + "-" + domCode + "-";

  // Find the highest sequence number already used with this prefix
  const allData = sheet.getDataRange().getValues();
  let maxSeq = 0;

  for (let i = 1; i < allData.length; i++) { // skip header row
    const existingId = allData[i][COL_IDEA.ID - 1];
    if (existingId && existingId.toString().startsWith(prefix)) {
      const parts = existingId.toString().split("-");
      const seq   = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
    }
  }

  // Also check other sheets so IDs are globally unique
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToCheck = [
    SHEET.MASTER, SHEET.RESEARCH, SHEET.SCRIPT,
    SHEET.VISUAL, SHEET.PUBLISHING
  ];

  sheetsToCheck.forEach(sheetName => {
    try {
      const s = ss.getSheetByName(sheetName);
      if (!s) return;
      const data = s.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const id = data[i][0]; // ID is always column A
        if (id && id.toString().startsWith(prefix)) {
          const parts = id.toString().split("-");
          const seq   = parseInt(parts[parts.length - 1], 10);
          if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
        }
      }
    } catch (e) {
      // Sheet might not exist yet — skip silently
    }
  });

  const newSeq = String(maxSeq + 1).padStart(3, "0"); // "001"
  return prefix + newSeq; // "GX-2605-TECH-001"
}

// ── Manual ID generation (menu fallback) ─────────────────────────────────────
function generateIdForSelectedRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEET.IDEA) {
    SpreadsheetApp.getUi().alert(
      "Please select a row in the Idea Catalogue tab first."
    );
    return;
  }

  const row     = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Please select a data row (not the header).");
    return;
  }

  const company = sheet.getRange(row, COL_IDEA.COMPANY).getValue();
  if (!company || company.toString().trim() === "") {
    SpreadsheetApp.getUi().alert(
      "Please fill in the Company column before generating an ID."
    );
    return;
  }

  const existingId = sheet.getRange(row, COL_IDEA.ID).getValue();
  if (existingId && existingId.toString().trim() !== "") {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "ID Already Exists",
      "This row already has ID: " + existingId + "\nDo you want to regenerate it?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  const newId = generateId(sheet, row);
  sheet.getRange(row, COL_IDEA.ID).setValue(newId);

  SpreadsheetApp.getUi().alert("ID generated: " + newId);
}
