/* ============================================================================
   Company_Selector_Skill.gs — GovernX Content OS
   Stage 0 — Strategic Company Selection

   ARCHITECTURE (v3 — Sheet-Driven Database):
   The company database lives in a dedicated "Company Database" tab
   in your Google Sheet — not embedded in this code.

   You can add, edit, or remove companies at any time without touching code.
   The script reads from that tab on every run.

   DATABASE TAB COLUMNS (Company Database):
   A: Company        B: Country       C: Domain        D: Industry
   E: Type           F: Loss/Gain $B  G: Hook          H: GRC/BPR Angle
   I: Confidence     J: Source URL    K: Suggested ✓   (checkbox)

   HOW DEDUPLICATION WORKS:
   Column K (Suggested) is a checkbox.
   When Stage 0 adds a company to the Idea Catalogue → it ticks column K.
   Stage 0 only reads rows where column K is UNCHECKED.
   You can manually uncheck a row to allow it to be suggested again.

   SETUP:
   Run "🗄️ Setup Company Database Tab" from the GovernX menu ONCE.
   This creates the tab and pre-populates it with all 200 cases.
   After that, manage the tab manually — add rows, edit hooks, add new companies.
   ============================================================================ */


// ── Sheet tab names — your actual tab names ──────────────────────────────────
const COMPANY_DB_SUCCESS = "DB_S";
const COMPANY_DB_FAILURE = "DB_F";

// ── Column indices — Company Database_SUCCESS tab (1-based) ──────────────────
// Rank | Company | Country | Domain | Industry | Field |
// Governance Inflection Date | Estimated Gain USD B | Gain Basis |
// Governance Strength Angle | Video Hook | Source URL |
// Research Confidence | Notes | Check
const COL_DB_SUCCESS = {
  RANK       : 1,   // A
  COMPANY    : 2,   // B
  COUNTRY    : 3,   // C
  DOMAIN     : 4,   // D
  INDUSTRY   : 5,   // E
  FIELD      : 6,   // F
  DATE       : 7,   // G — Governance Inflection Date / Period
  VALUE_B    : 8,   // H — Estimated Governance-Enabled Gain USD B
  GAIN_BASIS : 9,   // I
  GRC_ANGLE  : 10,  // J — Governance Strength Angle
  HOOK       : 11,  // K — Video Reverse-Engineering Hook
  SOURCE_URL : 12,  // L
  CONFIDENCE : 13,  // M — Research Confidence
  NOTES      : 14,  // N
  CHECK      : 15   // O — checkbox: TRUE = already suggested
};

// ── Column indices — Company Database_FAILURE tab (1-based) ──────────────────
// Rank | Company | Country | Domain | Industry | Field |
// Collapse/Trigger Date | Estimated Loss USD M | Estimated Loss USD B |
// Loss Basis | Governance Failure Angle | Video Hook | Source URL |
// Research Confidence | Notes | Check
const COL_DB_FAILURE = {
  RANK       : 1,   // A
  COMPANY    : 2,   // B
  COUNTRY    : 3,   // C
  DOMAIN     : 4,   // D
  INDUSTRY   : 5,   // E
  FIELD      : 6,   // F
  DATE       : 7,   // G — Collapse/Trigger Date
  LOSS_M     : 8,   // H — Estimated Loss USD M
  VALUE_B    : 9,   // I — Estimated Loss USD B
  LOSS_BASIS : 10,  // J
  GRC_ANGLE  : 11,  // K — Governance Failure Angle
  HOOK       : 12,  // L — Video Reverse-Engineering Hook
  SOURCE_URL : 13,  // M
  CONFIDENCE : 14,  // N — Research Confidence
  NOTES      : 15,  // O
  CHECK      : 16   // P — checkbox: TRUE = already suggested
};


// ══════════════════════════════════════════════════════════════════════════════
// SETUP — Validates your existing database tabs and ensures Check column
// has proper checkboxes. Run once after adding the script.
// Your tabs already exist — this does NOT overwrite them.
// ══════════════════════════════════════════════════════════════════════════════
function setupCompanyDatabase() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const successSheet = ss.getSheetByName(COMPANY_DB_SUCCESS);
  const failureSheet = ss.getSheetByName(COMPANY_DB_FAILURE);

  const issues = [];
  if (!successSheet) issues.push("Tab not found: " + COMPANY_DB_SUCCESS);
  if (!failureSheet) issues.push("Tab not found: " + COMPANY_DB_FAILURE);

  if (issues.length > 0) {
    ui.alert("⚠️ Database Tabs Not Found",
      issues.join("\n") + "\n\nExpected tab names:\n  • " +
      COMPANY_DB_SUCCESS + "\n  • " + COMPANY_DB_FAILURE +
      "\n\nCheck that tab names match exactly (case-sensitive).",
      ui.ButtonSet.OK);
    return;
  }

  const successRows = successSheet.getLastRow() - 1;
  const failureRows = failureSheet.getLastRow() - 1;

  if (successRows > 0)
    successSheet.getRange(2, COL_DB_SUCCESS.CHECK, successRows, 1).insertCheckboxes();
  if (failureRows > 0)
    failureSheet.getRange(2, COL_DB_FAILURE.CHECK, failureRows, 1).insertCheckboxes();

  SpreadsheetApp.flush();

  ui.alert("✅ Database Validated",
    "Both tabs found:\n\n" +
    "• " + COMPANY_DB_SUCCESS + ": " + successRows + " companies\n" +
    "• " + COMPANY_DB_FAILURE + ": " + failureRows + " companies\n\n" +
    "Check column (last col) = deduplication checkbox.\n" +
    "Stage 0 ticks it automatically when a company is added.\n" +
    "Untick manually to allow re-suggestion.\n\n" +
    "Run Stage 0 from the GovernX menu.",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// READ COMPANY DATABASE FROM SHEET
// Reads from Company Database_Success and Company Database_Failure tabs
// Returns only rows where Check = FALSE (not yet suggested)
// ══════════════════════════════════════════════════════════════════════════════
function readCompanyDatabase(ss, typeFilter) {
  const successSheet = ss.getSheetByName(COMPANY_DB_SUCCESS);
  const failureSheet = ss.getSheetByName(COMPANY_DB_FAILURE);

  if (!successSheet || !failureSheet) {
    throw new Error(
      "Database tabs not found.\n\nExpected: '" + COMPANY_DB_SUCCESS +
      "' and '" + COMPANY_DB_FAILURE + "'\n" +
      "Run '🗄️ Setup Company Database Tab' from the menu to validate."
    );
  }

  const companies = [];

  // ── Read SUCCESS tab ────────────────────────────────────────────────────────
  if (typeFilter !== "1") {
    const sData = successSheet.getDataRange().getValues();
    for (let i = 1; i < sData.length; i++) {
      const row     = sData[i];
      const company = row[COL_DB_SUCCESS.COMPANY - 1].toString().trim();
      const checked = row[COL_DB_SUCCESS.CHECK   - 1];
      if (!company || checked === true) continue;
      companies.push({
        sheetName  : COMPANY_DB_SUCCESS,
        sheetRow   : i + 1,
        type       : "Success",
        company    : company,
        country    : row[COL_DB_SUCCESS.COUNTRY    - 1].toString().trim(),
        domain     : row[COL_DB_SUCCESS.DOMAIN     - 1].toString().trim(),
        industry   : row[COL_DB_SUCCESS.INDUSTRY   - 1].toString().trim(),
        field      : row[COL_DB_SUCCESS.FIELD      - 1].toString().trim(),
        date       : row[COL_DB_SUCCESS.DATE       - 1].toString().trim(),
        valueB     : parseFloat(row[COL_DB_SUCCESS.VALUE_B    - 1]) || 0,
        hook       : row[COL_DB_SUCCESS.HOOK       - 1].toString().trim(),
        angle      : row[COL_DB_SUCCESS.GRC_ANGLE  - 1].toString().trim(),
        confidence : row[COL_DB_SUCCESS.CONFIDENCE - 1].toString().trim(),
        sourceUrl  : row[COL_DB_SUCCESS.SOURCE_URL - 1].toString().trim()
      });
    }
  }

  // ── Read FAILURE tab ────────────────────────────────────────────────────────
  if (typeFilter !== "2") {
    const fData = failureSheet.getDataRange().getValues();
    for (let i = 1; i < fData.length; i++) {
      const row     = fData[i];
      const company = row[COL_DB_FAILURE.COMPANY - 1].toString().trim();
      const checked = row[COL_DB_FAILURE.CHECK   - 1];
      if (!company || checked === true) continue;
      companies.push({
        sheetName  : COMPANY_DB_FAILURE,
        sheetRow   : i + 1,
        type       : "Collapse",
        company    : company,
        country    : row[COL_DB_FAILURE.COUNTRY    - 1].toString().trim(),
        domain     : row[COL_DB_FAILURE.DOMAIN     - 1].toString().trim(),
        industry   : row[COL_DB_FAILURE.INDUSTRY   - 1].toString().trim(),
        field      : row[COL_DB_FAILURE.FIELD      - 1].toString().trim(),
        date       : row[COL_DB_FAILURE.DATE       - 1].toString().trim(),
        valueB     : parseFloat(row[COL_DB_FAILURE.VALUE_B    - 1]) || 0,
        hook       : row[COL_DB_FAILURE.HOOK       - 1].toString().trim(),
        angle      : row[COL_DB_FAILURE.GRC_ANGLE  - 1].toString().trim(),
        confidence : row[COL_DB_FAILURE.CONFIDENCE - 1].toString().trim(),
        sourceUrl  : row[COL_DB_FAILURE.SOURCE_URL - 1].toString().trim()
      });
    }
  }

  return companies;
}


// ══════════════════════════════════════════════════════════════════════════════
// TICK "SUGGESTED" CHECKBOX FOR A COMPANY
// Called after a company is added to the Idea Catalogue
// ══════════════════════════════════════════════════════════════════════════════
function markCompanyAsSuggested(ss, companyName) {
  // Search both database tabs for the company and tick its Check checkbox
  try {
    const tabs = [
      { sheet: ss.getSheetByName(COMPANY_DB_SUCCESS), col: COL_DB_SUCCESS.CHECK, companyCol: COL_DB_SUCCESS.COMPANY },
      { sheet: ss.getSheetByName(COMPANY_DB_FAILURE), col: COL_DB_FAILURE.CHECK, companyCol: COL_DB_FAILURE.COMPANY }
    ];

    for (const tab of tabs) {
      if (!tab.sheet) continue;
      const data = tab.sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][tab.companyCol - 1].toString().trim().toLowerCase() ===
            companyName.toLowerCase()) {
          tab.sheet.getRange(i + 1, tab.col).setValue(true);
          Logger.log("✅ Marked as suggested: " + companyName +
            " in " + tab.sheet.getName() + " row " + (i + 1));
          return;
        }
      }
    }
    Logger.log("Company not found in either database tab: " + companyName);
  } catch (e) {
    Logger.log("markCompanyAsSuggested error (non-fatal): " + e.message);
  }
}


const SELECTOR_SYSTEM_CONTEXT = `
You are the Content Strategy Director for GovernX — a YouTube channel targeting 
Arabic-speaking business leaders, C-Suite executives, and board members across 
the Gulf and Arab world.

YOUR ROLE: Select the next company case(s) to produce — not based on instinct, 
but on a scored, strategic framework.

You never guess. You score, rank, and justify.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCORING MODEL — FIVE CRITERIA (20 points each = 100 total)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CRITERION 1 — STORY STRENGTH (20 pts)
Does the governance angle produce a CLEAR, DRAMATIC reverse-engineering arc?
- Can you name the specific governance failure/success in one sentence? (+5)
- Is there a single identifiable decision moment that changed everything? (+5)
- Does the root cause connect directly to a GRC or BPR principle? (+5)
- Is the story resolvable — does it have a clear lesson, not just a narrative? (+5)

CRITERION 2 — AUDIENCE RESONANCE (20 pts)
How relevant is this to Gulf/Arab C-Suite executives and board members?
- Is the industry directly present or growing in the GCC/Arab world? (+5)
- Does the failure/success type mirror governance challenges in the region? (+5)
- Is the scale of the outcome ($B loss/gain) credible to a C-Suite audience? (+5)
- Does the case have a cultural or geographic proximity advantage? (+5)
  (MENA company = maximum +5 | Emerging market = +3 | Western = +1)

CRITERION 3 — CONTRAST PAIR AVAILABILITY (20 pts)
Can this case be paired with a strong mirror case in the same industry?
- Is there a clear collapse↔success mirror in the same domain/industry? (+10)
- Does the mirror case make the GRC lesson sharper by comparison? (+5)
- Is the contrast pair strong enough to anchor a two-episode series? (+5)
  (No pair available = 0 pts)

CRITERION 4 — PRODUCTION READINESS (20 pts)
How ready is this case for immediate production?
- Research Confidence = High: +15 | Medium: +8 | Low: +2
- Are primary sources clearly available and accessible? (+5 if yes)

CRITERION 5 — CALENDAR DIVERSITY (20 pts)
Does selecting this case improve the content calendar?
- Domain NOT covered in last 3 videos: +8
- Country/region NOT covered in last 3 videos: +7
- Case type (collapse vs success) balances recent history: +5

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMAT — FOLLOW EXACTLY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Return a structured selection brief in the exact format specified in the prompt.
Be specific. Score every criterion with a number, not a description.
Name the contrast pair explicitly — don't just suggest a category.
`;


// ══════════════════════════════════════════════════════════════════════════════
// STAGE 0 — RUN COMPANY SELECTOR
// Called from GovernX menu before Stage 1
// ══════════════════════════════════════════════════════════════════════════════
function runCompanySelector() {

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Ask user for content type preference ──────────────────────────────────
  const typeResponse = ui.prompt(
    "Stage 0 — Company Selector  (1 of 5)",
    "What type of video are you planning next?\n\n" +
    "  1 — Collapse story (governance failure)\n" +
    "  2 — Success story (governance win)\n" +
    "  3 — Contrast pair (collapse + success mirror)\n" +
    "  4 — Surprise me (highest scoring regardless of type)\n\n" +
    "Enter 1, 2, 3, or 4:",
    ui.ButtonSet.OK_CANCEL
  );
  if (typeResponse.getSelectedButton() === ui.Button.CANCEL) return;
  const typeChoice = typeResponse.getResponseText().trim();

  // ── Domain filter ─────────────────────────────────────────────────────────
  const domainResponse = ui.prompt(
    "Stage 0 — Domain Filter  (2 of 5)",
    "Prefer a specific domain? Leave blank for no filter.\n\n" +
    "Examples: Finance, Technology, Automotive, Energy, Healthcare\n" +
    "MENA (to prioritize Gulf/Arab region cases)\n",
    ui.ButtonSet.OK_CANCEL
  );
  if (domainResponse.getSelectedButton() === ui.Button.CANCEL) return;
  const domainFilter = domainResponse.getResponseText().trim();

  // ── Language ──────────────────────────────────────────────────────────────
  const langResponse = ui.prompt(
    "Stage 0 — Language  (3 of 5)",
    "What language should this video be produced in?\n\n" +
    "  1 — English\n" +
    "  2 — Arabic\n" +
    "  3 — Bilingual  (Arabic voiceover + English structure)\n\n" +
    "Enter 1, 2, or 3:",
    ui.ButtonSet.OK_CANCEL
  );
  if (langResponse.getSelectedButton() === ui.Button.CANCEL) return;
  const langChoice = langResponse.getResponseText().trim();
  const languageFlag = { "1": "English", "2": "Arabic", "3": "Bilingual" }[langChoice] || "English";

  // ── Target Format ─────────────────────────────────────────────────────────
  const formatResponse = ui.prompt(
    "Stage 0 — Target Format  (4 of 5)",
    "What video format are you targeting?\n\n" +
    "  1 — Short (< 90s)         → 150–250 words\n" +
    "  2 — Standard (4–7 min)    → 600–900 words\n" +
    "  3 — Deep Dive (10–15 min) → 1500–2000 words\n\n" +
    "Enter 1, 2, or 3:",
    ui.ButtonSet.OK_CANCEL
  );
  if (formatResponse.getSelectedButton() === ui.Button.CANCEL) return;
  const formatChoice = formatResponse.getResponseText().trim();
  const targetFormat = {
    "1": "Short (< 90s)",
    "2": "Standard (4–7 min)",
    "3": "Deep Dive (10–15 min)"
  }[formatChoice] || "Standard (4–7 min)";

  // ── Series ────────────────────────────────────────────────────────────────
  const seriesResponse = ui.prompt(
    "Stage 0 — Series  (5 of 5)",
    "Which series should this video belong to?\n\n" +
    "  1 — Governance Collapse Series\n" +
    "  2 — BPR Turning Points\n" +
    "  3 — Risk Blind Spots\n" +
    "  4 — Leadership Decisions\n" +
    "  5 — System Design Failures\n" +
    "  0 — Let Claude decide based on the story\n\n" +
    "Enter 0–5:",
    ui.ButtonSet.OK_CANCEL
  );
  if (seriesResponse.getSelectedButton() === ui.Button.CANCEL) return;
  const seriesChoice = seriesResponse.getResponseText().trim();
  const seriesValue = {
    "1": "Governance Collapse Series",
    "2": "BPR Turning Points",
    "3": "Risk Blind Spots",
    "4": "Leadership Decisions",
    "5": "System Design Failures",
    "0": ""
  }[seriesChoice] !== undefined
    ? ({ "1": "Governance Collapse Series", "2": "BPR Turning Points",
         "3": "Risk Blind Spots", "4": "Leadership Decisions",
         "5": "System Design Failures", "0": "" }[seriesChoice])
    : "";

  // ── Read available companies — cap at 60 highest-confidence entries ────────
  // Sending the full database (~150+ rows) to Claude adds 3–5k tokens and
  // 15–20 extra seconds. Top-60 by confidence produces identical recommendations.
  let availableCompanies;
  try {
    availableCompanies = readCompanyDatabase(ss, typeChoice);
  } catch (dbErr) {
    ui.alert("⚠️ Database Not Found", dbErr.message, ui.ButtonSet.OK);
    return;
  }

  if (availableCompanies.length === 0) {
    ui.alert(
      "No Companies Available",
      "All companies in the database have already been suggested,\n" +
      "or no companies match your filter.\n\n" +
      "To reset: open the DB_S or DB_F tab and uncheck\n" +
      "the Check boxes for companies you want to reconsider.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Sort by confidence (High → Medium → Low), then cap
  const DB_CAP = 60;
  const confidenceRank = { "High": 3, "Medium": 2, "Low": 1, "": 0 };
  const sortedCompanies = availableCompanies
    .slice()
    .sort(function(a, b) {
      return (confidenceRank[b.confidence] || 0) - (confidenceRank[a.confidence] || 0);
    });
  const cappedCompanies  = sortedCompanies.slice(0, DB_CAP);
  const trimmedCount     = availableCompanies.length - cappedCompanies.length;

  Logger.log("Stage 0: " + availableCompanies.length + " available → " +
    cappedCompanies.length + " sent to Claude" +
    (trimmedCount > 0 ? " (" + trimmedCount + " low-confidence trimmed)" : ""));

  const excludedCompanies = [];

  ui.alert(
    "Stage 0 — Analyzing Cases",
    "Analyzing top " + cappedCompanies.length + " candidates\n" +
    "Language: " + languageFlag + " | Format: " + targetFormat +
    (seriesValue ? " | Series: " + seriesValue : "") + "\n\n" +
    "This takes about 15–20 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildSelectorPrompt(typeChoice, domainFilter, getRecentContentHistory(ss, 6), excludedCompanies, cappedCompanies);
    const raw    = callClaudeAsSelector(prompt);

    const selection = parseSelectorOutput(raw);
    const reportUrl = exportSelectionReport(selection, raw, getRecentContentHistory(ss, 6));

    const topPick    = selection.topPick;
    const topCompany = topPick ? topPick.company : "See report";
    const topScore   = topPick ? topPick.totalScore : "—";

    const confirmResponse = ui.alert(
      "✅ Company Selection Complete",
      "TOP RECOMMENDATION: " + topCompany + " (Score: " + topScore + "/100)\n\n" +
      (topPick ? "Discipline: " + topPick.discipline + "\n" : "") +
      (topPick ? "Type: " + topPick.type + "\n\n" : "\n") +
      (selection.contrastPair
        ? "CONTRAST PAIR: " + selection.contrastPair.collapse + " ↔ " + selection.contrastPair.success + "\n\n"
        : "") +
      "Will write to Ideas Bank with:\n" +
      "  Language   : " + languageFlag + "\n" +
      "  Format     : " + targetFormat + "\n" +
      (seriesValue ? "  Series     : " + seriesValue + "\n" : "  Series     : Claude will decide\n") +
      "\nFull Selection Report saved to Drive.\n\n" +
      "Add this company to your Ideas Bank now?",
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse === ui.Button.YES && topPick) {
      writeSelectionToIdeasBank(ss, topPick, selection, languageFlag, targetFormat, seriesValue);
      ui.alert(
        "✅ Added to Ideas Bank",
        topPick.company + " has been added to your Ideas Bank.\n" +
        "Select that row and run Stage 1 to begin production.",
        ui.ButtonSet.OK
      );
    }

  } catch (err) {
    logError("Stage 0 — Company Selector", "—", "Selector API Error", err.message);
    ui.alert("❌ Stage 0 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// BUILD SELECTOR PROMPT
// Sends full database + recent history + scoring model to Claude
// ══════════════════════════════════════════════════════════════════════════════
function buildSelectorPrompt(typeChoice, domainFilter, recentHistory, excludedCompanies, availableCompanies) {

  const typeInstruction = {
    "1": "Score and rank COLLAPSE cases only. Still identify the best contrast pair for each top pick.",
    "2": "Score and rank SUCCESS cases only. Still identify the best contrast pair for each top pick.",
    "3": "Score and rank CONTRAST PAIRS specifically — pairs where one collapse + one success share the same industry/domain. The pair score is the combined average.",
    "4": "Score ALL cases (collapse + success). Return the highest scoring regardless of type."
  }[typeChoice] || "Score ALL cases and return the highest scoring regardless of type.";

  const domainInstruction = domainFilter
    ? (domainFilter.toUpperCase() === "MENA"
        ? "PRIORITY FILTER: Apply a +5 bonus to any case from UAE, Saudi Arabia, Qatar, Egypt, Jordan, Kuwait, or other Arab/Gulf region companies."
        : "DOMAIN FILTER: Give extra consideration to the domain '" + domainFilter + "' — if two cases score equally, prefer this domain.")
    : "No domain filter — score all domains equally.";

  // Build compact DB representation from sheet data (passed in as availableCompanies)
  // Note: availableCompanies is read from the Company Database sheet tab
  // and already filtered — only unsugested companies are included
  const collapseList = availableCompanies
    .filter(c => c.type === "Collapse")
    .map((c, i) => `C${i+1}|${c.company}|${c.country}|${c.domain}|${c.industry}|$${c.valueB}B|${c.confidence}|"${c.hook}"`)
    .join("\n") || "No collapse cases available";

  const successList = availableCompanies
    .filter(c => c.type === "Success")
    .map((s, i) => `S${i+1}|${s.company}|${s.country}|${s.domain}|${s.industry}|$${s.valueB}B|${s.confidence}|"${s.hook}"`)
    .join("\n") || "No success cases available";

  const recentList = recentHistory.length > 0
    ? recentHistory.map(h => `  • ${h.company} (${h.domain}) — ${h.type} — ${h.date}`).join("\n")
    : "  • No previous content IDs found — this is the first batch.";

  // Build exclusion list for the prompt
  const excludedList = excludedCompanies && excludedCompanies.length > 0
    ? excludedCompanies.map(c => `  • ${c}`).join("\n")
    : "  • None — all companies are eligible.";

  return `
You are selecting the next GovernX video topic(s).

SCORING INSTRUCTION: ${typeInstruction}
${domainInstruction}

═══════════════════════════════════════════════════
EXCLUDED COMPANIES — DO NOT SUGGEST THESE
═══════════════════════════════════════════════════
The following companies are already in the Idea Catalogue or have been
produced. NEVER suggest any of these — they are permanently excluded.
${excludedList}

═══════════════════════════════════════════════════
RECENT CONTENT HISTORY (avoid repetition)
═══════════════════════════════════════════════════
${recentList}

═══════════════════════════════════════════════════
COLLAPSE DATABASE (format: ID|Company|Country|Domain|Industry|Loss|Confidence|Hook)
═══════════════════════════════════════════════════
${collapseList}

═══════════════════════════════════════════════════
SUCCESS DATABASE (format: ID|Company|Country|Domain|Industry|Gain|Confidence|Hook)
═══════════════════════════════════════════════════
${successList}

═══════════════════════════════════════════════════
YOUR OUTPUT — RETURN IN EXACTLY THIS FORMAT
═══════════════════════════════════════════════════

SELECTION_REPORT_START

TOP_PICKS_START
[Return top 5 candidates, ranked by total score. Use one CANDIDATE block per pick.]

CANDIDATE_START
RANK: [1–5]
COMPANY: [exact company name from database]
TYPE: [Collapse | Success]
COUNTRY: [country]
DOMAIN: [domain]
INDUSTRY: [industry]
SCALE: [$XB loss/gain]
HOOK: [hook from database]
TOTAL_SCORE: [number out of 100]
SCORE_BREAKDOWN:
  STORY_STRENGTH: [X/20] — [1 sentence why]
  AUDIENCE_RESONANCE: [X/20] — [1 sentence why]
  CONTRAST_PAIR_AVAILABILITY: [X/20] — [1 sentence why — name the pair]
  PRODUCTION_READINESS: [X/20] — [confidence + sources note]
  CALENDAR_DIVERSITY: [X/20] — [1 sentence why]
GRC_BPR_ANGLE: [The specific GRC or BPR principle this case teaches — be precise]
DISCIPLINE: [GRC | BPR | Both]
RECOMMENDED_HOOK_AR: [Suggest a sharpened Arabic-market version of the hook in English — do NOT write Arabic text here]
CANDIDATE_END

[repeat for ranks 2–5]
TOP_PICKS_END

─────────────────────────────────────────
CONTRAST_PAIR_START
RECOMMENDED_PAIR: [YES | NO]
COLLAPSE_COMPANY: [company name | N/A]
SUCCESS_COMPANY: [company name | N/A]
SHARED_DOMAIN: [domain they share]
PAIR_RATIONALE: [2–3 sentences: why this comparison sharpens the GRC lesson. What does seeing both sides teach that one case alone cannot?]
PAIR_EPISODE_FORMAT: [Suggest how to structure this as a 2-episode series: Episode 1 topic, Episode 2 topic]
CONTRAST_PAIR_END

─────────────────────────────────────────
CALENDAR_SUGGESTION_START
[Suggest a 4-video sequence based on the top picks and contrast pair]
VIDEO_1: [Company | Type | Domain | Why this goes first]
VIDEO_2: [Company | Type | Domain | Why this is the follow-up]
VIDEO_3: [Company | Type | Domain | Strategic reason]
VIDEO_4: [Company | Type | Domain | Strategic reason]
CALENDAR_SUGGESTION_END

─────────────────────────────────────────
GEOGRAPHIC_DIVERSITY_NOTE: [1–2 sentences flagging if the shortlist is over-indexed on one region, and which MENA/emerging market cases should be considered sooner]

SELECTOR_VERDICT: [1–2 sentences: your final recommendation — which company to produce next and why in one clear statement]

SELECTION_REPORT_END
`;
}


// ══════════════════════════════════════════════════════════════════════════════
// CALL CLAUDE AS SELECTOR
// Uses SELECTOR_SYSTEM_CONTEXT — separate from all other Claude instances
// ══════════════════════════════════════════════════════════════════════════════
function callClaudeAsSelector(finalPrompt) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("ANTHROPIC_API_KEY");

  if (!apiKey) throw new Error("ANTHROPIC_API_KEY missing from Script Properties");

  const MAX_TRIES     = 3;
  const RETRY_WAIT_MS = 8000;
  const RATELIMIT_MS  = 20000;
  let   lastError     = "";

  for (let attempt = 1; attempt <= MAX_TRIES; attempt++) {
    try {
      const payload = {
        model      : ANTHROPIC_MODEL,
        max_tokens : 8000,
        system     : SELECTOR_SYSTEM_CONTEXT,
        messages   : [{ role: "user", content: finalPrompt }]
      };

      const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
        method            : "post",
        contentType       : "application/json",
        headers: {
          "x-api-key"         : apiKey,
          "anthropic-version" : "2023-06-01",
          "accept"            : "application/json"
        },
        payload           : JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      const body = response.getContentText();

      if (code === 200) {
        const json = JSON.parse(body);
        return json.content[0].text;
      }

      // Claude credit exhausted → free Groq fallback so the company selector keeps flowing.
      if (code === 400 && /credit balance is too low/i.test(body)) {
        const fb = groqFallback_(SELECTOR_SYSTEM_CONTEXT, finalPrompt, 8000, "stage_0_selector");
        if (fb !== null) return fb;
        throw new Error("Claude credit is zero and no GROQ_API_KEY is set. Add Anthropic credits, " +
          "or set GROQ_API_KEY in Script Properties to keep producing on the free fallback.");
      }

      if (code === 529 || code === 429) {
        lastError = "API error " + code;
        Utilities.sleep(code === 429 ? RATELIMIT_MS : RETRY_WAIT_MS);
        continue;
      }

      throw new Error("Selector API error " + code + ": " + body);

    } catch (err) {
      lastError = err.message;
      if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
    }
  }

  throw new Error("Selector API failed after " + MAX_TRIES + " attempts. Last: " + lastError);
}


// ══════════════════════════════════════════════════════════════════════════════
// PARSE SELECTOR OUTPUT
// ══════════════════════════════════════════════════════════════════════════════
function parseSelectorOutput(raw) {

  const result = {
    topPick      : null,
    candidates   : [],
    contrastPair : null,
    calendar     : [],
    geoNote      : "",
    verdict      : ""
  };

  // ── Parse candidates ────────────────────────────────────────────────────────
  const candidateBlocks = raw.match(/CANDIDATE_START([\s\S]*?)CANDIDATE_END/g) || [];
  candidateBlocks.forEach(block => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };

    const candidate = {
      rank         : get("RANK"),
      company      : get("COMPANY"),
      type         : get("TYPE"),
      country      : get("COUNTRY"),
      domain       : get("DOMAIN"),
      industry     : get("INDUSTRY"),
      scale        : get("SCALE"),
      hook         : get("HOOK"),
      totalScore   : get("TOTAL_SCORE"),
      grcAngle     : get("GRC_BPR_ANGLE"),
      discipline   : get("DISCIPLINE"),
      arabicHook   : get("RECOMMENDED_HOOK_AR")
    };

    result.candidates.push(candidate);
    if (candidate.rank === "1") result.topPick = candidate;
  });

  // ── Parse contrast pair ────────────────────────────────────────────────────
  const pairBlock = raw.match(/CONTRAST_PAIR_START([\s\S]*?)CONTRAST_PAIR_END/);
  if (pairBlock) {
    const get = (field) => {
      const m = pairBlock[1].match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };
    const getBlock = (field) => {
      const m = pairBlock[1].match(new RegExp(field + ":\\s*([\\s\\S]*?)(?=\\n[A-Z_]+:|CONTRAST_PAIR_END)"));
      return m ? m[1].trim() : "";
    };

    if (get("RECOMMENDED_PAIR") === "YES") {
      result.contrastPair = {
        collapse      : get("COLLAPSE_COMPANY"),
        success       : get("SUCCESS_COMPANY"),
        domain        : get("SHARED_DOMAIN"),
        rationale     : getBlock("PAIR_RATIONALE"),
        episodeFormat : getBlock("PAIR_EPISODE_FORMAT")
      };
    }
  }

  // ── Parse calendar ────────────────────────────────────────────────────────
  const calBlock = raw.match(/CALENDAR_SUGGESTION_START([\s\S]*?)CALENDAR_SUGGESTION_END/);
  if (calBlock) {
    for (let i = 1; i <= 4; i++) {
      const m = calBlock[1].match(new RegExp("VIDEO_" + i + ":\\s*(.+)"));
      if (m) result.calendar.push(m[1].trim());
    }
  }

  // ── Geographic note + verdict ──────────────────────────────────────────────
  const geoMatch = raw.match(/GEOGRAPHIC_DIVERSITY_NOTE:\s*(.+)/);
  if (geoMatch) result.geoNote = geoMatch[1].trim();

  const verdictMatch = raw.match(/SELECTOR_VERDICT:\s*(.+)/);
  if (verdictMatch) result.verdict = verdictMatch[1].trim();

  return result;
}


// ══════════════════════════════════════════════════════════════════════════════
// EXPORT SELECTION REPORT TO GOOGLE DRIVE
// ══════════════════════════════════════════════════════════════════════════════
function exportSelectionReport(selection, rawOutput, recentHistory) {

  const divider = "═".repeat(60);
  const line    = "─".repeat(60);
  const now     = new Date().toLocaleString();

  let doc = "";
  doc += divider + "\n";
  doc += "GOVERNX — COMPANY SELECTION REPORT\n";
  doc += divider + "\n";
  doc += "Generated: " + now + "\n";
  doc += "Candidates Evaluated: " + (rawOutput ? "See shortlist below" : "—") + "\n";
  doc += divider + "\n\n";

  // ── Recent history ─────────────────────────────────────────────────────────
  doc += "RECENT CONTENT HISTORY\n" + line + "\n";
  if (recentHistory.length > 0) {
    recentHistory.forEach(h => {
      doc += "  • " + h.company + " (" + h.domain + ") — " + h.date + "\n";
    });
  } else {
    doc += "  No previous content found.\n";
  }
  doc += "\n";

  // ── Top picks ──────────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "TOP 5 CANDIDATES — RANKED BY SCORE\n";
  doc += line + "\n\n";

  selection.candidates.forEach((c, idx) => {
    const icon = idx === 0 ? "🥇" : idx === 1 ? "🥈" : idx === 2 ? "🥉" : "  ";
    doc += icon + " RANK " + c.rank + ": " + c.company + " (" + c.type + ")\n";
    doc += "  Score       : " + c.totalScore + "/100\n";
    doc += "  Country     : " + c.country + "\n";
    doc += "  Domain      : " + c.domain + " / " + c.industry + "\n";
    doc += "  Scale       : " + c.scale + "\n";
    doc += "  GRC Angle   : " + c.grcAngle + "\n";
    doc += "  Discipline  : " + c.discipline + "\n";
    doc += "  Hook        : " + c.hook + "\n";
    doc += "  Arabic Hook : " + c.arabicHook + "\n";
    doc += "\n";
  });

  // ── Contrast pair ──────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "RECOMMENDED CONTRAST PAIR\n";
  doc += line + "\n\n";

  if (selection.contrastPair) {
    doc += "⚡ COLLAPSE: " + selection.contrastPair.collapse + "\n";
    doc += "✅ SUCCESS : " + selection.contrastPair.success  + "\n";
    doc += "Domain     : " + selection.contrastPair.domain   + "\n\n";
    doc += "WHY THIS PAIR:\n" + selection.contrastPair.rationale + "\n\n";
    doc += "EPISODE FORMAT:\n" + selection.contrastPair.episodeFormat + "\n\n";
  } else {
    doc += "No strong contrast pair identified for this selection.\n\n";
  }

  // ── Content calendar ───────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "4-VIDEO CONTENT CALENDAR SUGGESTION\n";
  doc += line + "\n\n";

  selection.calendar.forEach((entry, idx) => {
    doc += "Video " + (idx + 1) + ": " + entry + "\n";
  });
  doc += "\n";

  // ── Geographic note ────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "GEOGRAPHIC DIVERSITY NOTE\n";
  doc += line + "\n\n";
  doc += selection.geoNote + "\n\n";

  // ── Final verdict ──────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SELECTOR VERDICT\n";
  doc += line + "\n\n";
  doc += selection.verdict + "\n\n";
  doc += divider + "\n";
  doc += "END OF SELECTION REPORT\n";
  doc += divider + "\n";

  // ── Save to Drive ──────────────────────────────────────────────────────────
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const rootFolder  = getOrCreateRootFolder();
  const title       = "GovernX — Selection Report — " + Utilities.formatDate(new Date(), "GMT+3", "yyyy-MM-dd HH:mm");

  const newDoc = DocumentApp.create(title);
  newDoc.getBody().setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: "Courier New" });
  newDoc.getBody().setText(doc);
  newDoc.saveAndClose();

  const file = DriveApp.getFileById(newDoc.getId());
  rootFolder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  Logger.log("Selection Report saved: " + newDoc.getUrl());
  return newDoc.getUrl();
}


// ══════════════════════════════════════════════════════════════════════════════
// WRITE TOP PICK TO IDEAS BANK
// Pre-fills company, domain, discipline, hook into the Ideas Bank sheet
// Ready for the user to select and run Stage 1
// ══════════════════════════════════════════════════════════════════════════════
function writeSelectionToIdeasBank(ss, topPick, selection, languageFlag, targetFormat, seriesValue) {

  const ideasSheet = ss.getSheetByName(SHEET.IDEA);
  if (!ideasSheet) {
    SpreadsheetApp.getUi().alert(
      "Ideas Bank tab not found. The selection report has been saved to Drive.\n" +
      "Add " + topPick.company + " to your Ideas Bank manually."
    );
    return;
  }

  // Find first truly empty row by scanning Company column
  let newRow = 2;
  const ideasData = ideasSheet.getDataRange().getValues();
  for (let i = 1; i < ideasData.length; i++) {
    const companyVal = ideasData[i][COL_IDEA.COMPANY - 1].toString().trim();
    if (companyVal === "") {
      newRow = i + 1;
      break;
    }
    newRow = i + 2;
  }

  // Contrast pair note for the Note column
  const pairNote = selection.contrastPair
    ? "Contrast pair: " + selection.contrastPair.collapse + " ↔ " + selection.contrastPair.success
    : "";

  // ── Clear data validations on all columns we're writing to ─────────────────
  [COL_IDEA.COMPANY, COL_IDEA.DOMAIN, COL_IDEA.INDUSTRY, COL_IDEA.FIELD,
   COL_IDEA.TYPE, COL_IDEA.INITIAL_ANGLE, COL_IDEA.LANGUAGE_FLAG,
   COL_IDEA.TARGET_FORMAT, COL_IDEA.SERIES, COL_IDEA.NOTE].forEach(col => {
    ideasSheet.getRange(newRow, col).clearDataValidations();
  });

  // ── Write all fields to the correct columns ───────────────────────────────
  ideasSheet.getRange(newRow, COL_IDEA.COMPANY      ).setValue(topPick.company);
  ideasSheet.getRange(newRow, COL_IDEA.DOMAIN       ).setValue(topPick.domain);
  ideasSheet.getRange(newRow, COL_IDEA.INDUSTRY     ).setValue(topPick.industry);
  ideasSheet.getRange(newRow, COL_IDEA.FIELD        ).setValue(topPick.industry);    // Field = sub-sector (same as industry from DB)
  ideasSheet.getRange(newRow, COL_IDEA.TYPE         ).setValue(topPick.type);
  ideasSheet.getRange(newRow, COL_IDEA.INITIAL_ANGLE).setValue(topPick.grcAngle);
  ideasSheet.getRange(newRow, COL_IDEA.LANGUAGE_FLAG).setValue(languageFlag);
  ideasSheet.getRange(newRow, COL_IDEA.TARGET_FORMAT).setValue(targetFormat);
  ideasSheet.getRange(newRow, COL_IDEA.SERIES       ).setValue(
    seriesValue || ""   // blank = Stage 1 lets Claude decide from taxonomy
  );
  ideasSheet.getRange(newRow, COL_IDEA.NOTE         ).setValue(
    "Selector Score: " + topPick.totalScore + "/100" +
    "\nDiscipline: " + topPick.discipline +
    (pairNote ? "\n" + pairNote : "") +
    "\n[SELECTOR-ADDED " + new Date().toLocaleDateString() + "]"
  );

  // Style the new row — light green = selector-approved
  ideasSheet.getRange(newRow, 1, 1, ideasSheet.getLastColumn())
    .setBackground("#E8F5E9");

  SpreadsheetApp.flush();
  Logger.log("Selection written to Ideas Bank: " + topPick.company +
    " | Lang: " + languageFlag +
    " | Format: " + targetFormat +
    " | Series: " + (seriesValue || "TBD"));

  // Tick the "Suggested ✓" checkbox in Company Database
  markCompanyAsSuggested(ss, topPick.company);
}


// ══════════════════════════════════════════════════════════════════════════════
// GET EXISTING IDEA CATALOGUE COMPANIES
// Reads ALL companies already in Idea Catalogue to build exclusion list
// Prevents same company being suggested again once it's been added
// ══════════════════════════════════════════════════════════════════════════════
function getExistingIdeaCatalogueCompanies(ss) {
  const excluded = [];
  try {
    const ideaSheet = ss.getSheetByName(SHEET.IDEA);
    if (!ideaSheet) return excluded;

    const data = ideaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const company = data[i][COL_IDEA.COMPANY - 1].toString().trim();
      if (company && company !== "") {
        excluded.push(company.toLowerCase());
      }
    }
  } catch (e) {
    Logger.log("Could not read Idea Catalogue for exclusions: " + e.message);
  }
  return excluded;
}


// ══════════════════════════════════════════════════════════════════════════════
// GET RECENT CONTENT HISTORY
// Reads last N entries from Master Content tab to inform diversity scoring
// ══════════════════════════════════════════════════════════════════════════════
function getRecentContentHistory(ss, limit) {

  const history = [];

  try {
    const masterSheet = ss.getSheetByName(SHEET.MASTER);
    if (!masterSheet) return history;

    const data     = masterSheet.getDataRange().getValues();
    const maxRows  = Math.min(data.length, limit + 1);

    for (let i = data.length - 1; i >= 1 && history.length < limit; i--) {
      const row = data[i];
      if (!row[COL_MASTER.ID - 1]) continue;

      history.push({
        id      : row[COL_MASTER.ID       - 1],
        company : row[COL_MASTER.TITLE    - 1] || "Unknown",
        domain  : row[COL_MASTER.DOMAIN   - 1] || "Unknown",
        type    : row[COL_MASTER.TYPE     - 1] || "Unknown",
        date    : ""
      });
    }
  } catch (e) {
    Logger.log("Could not read recent history: " + e.message);
  }

  return history;
}


// ══════════════════════════════════════════════════════════════════════════════
// GET OR CREATE ROOT FOLDER
// Creates/finds the GovernX root folder in Google Drive
// ══════════════════════════════════════════════════════════════════════════════
function getOrCreateRootFolder() {
  // Selector reports go into GovernX Production Packages root
  try {
    return DriveApp.getFolderById("1yErZa4vpGB-iAqetCgXKWepwQn1tMDSQ");
  } catch(e) {
    const s = DriveApp.getFoldersByName("GovernX Production Packages");
    return s.hasNext() ? s.next() : DriveApp.createFolder("GovernX Production Packages");
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// MENU INTEGRATION
// Add to Menu.gs in onOpen() BEFORE Stage 1:
//
//   .addItem("🗄️  Setup Company Database Tab",              "setupCompanyDatabase")
//   .addItem("0️⃣  Stage 0 — Company Selector",             "runCompanySelector")
//   .addSeparator()
//   .addItem("1️⃣  Stage 1 — Generate Master Content",      "generateMasterContent")
// ══════════════════════════════════════════════════════════════════════════════


// ══════════════════════════════════════════════════════════════════════════════ 