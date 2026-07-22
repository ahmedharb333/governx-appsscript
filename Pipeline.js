/* ============================================================================
   pipeline.gs — GovernX Content OS
   VERSION 2.1 — Data-driven pipeline upgrade

   UPGRADES IN v2.0:
   + Stage key arguments on all callClaude() calls (adaptive thinking per stage)
   + Stage 3 QA Critic: callClaudeAsEvaluator() replaces validateOutput for scripts
   + buildMasterContentPrompt: ROOT_CAUSE_ANALYST_SKILL injected
   + buildScriptPrompt: ROOT_CAUSE_ANALYST_SKILL + NARRATIVE_ARCHITECT_SKILL + ARABIC_VOICE_SKILL injected
   + buildScenesPrompt: VISUAL_INTELLIGENCE_SKILL injected
   UPGRADES IN v2.1:
   + DATA_RESEARCH_SKILL injected into buildResearchPrompt (Stage 2)
   + writeResearchDatabase() parses DATA_N blocks -> Data-type rows (light blue)
   + DATA_SCRIPT_SKILL injected into buildScriptPrompt (Stage 3)
   + getResearchSources() now reads NOTE column for Data payload
   (Skills 3-9 defined in Skills_Library.gs)

   Stage 1:  Master Content Table
   Stage 2:  Research Database
   Stage 3:  Script Bank (with QA Critic evaluation)
   Stage 4:  Visual Library
   Stage 4.5: Director Pass (Director_Skill.gs)
   Stage 5:  Publishing Tracker Row
   Stage 6:  Production Package
   Stage 7:  Voiceover Audio (ElevenLabs)
   Stage 8A: Scene Processing (Pexels + KlingAI)
   Stage 8B: Collect KlingAI Clips
   Stage 8C: Export Veo Prompts
   Stage 9:  Assembly Guide (CapCut)
   Stage 10: YouTube Metadata
   ============================================================================ */

// ── Shared: get active Idea Catalogue row data ────────────────────────────────
function getActiveIdeaRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== SHEET.IDEA) {
    SpreadsheetApp.getUi().alert(
      "Please select a row in the Idea Catalogue tab first."
    );
    return null;
  }

  const row = sheet.getActiveCell().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("Please select a data row (not the header).");
    return null;
  }

  // ── ID is generated here, never by hand ────────────────────────────────────
  // The onEdit trigger covers a human typing a Company, but it CANNOT cover a
  // programmatic write (simple triggers never fire on setValue), so a row created
  // by Stage 0 — Company Selector arrived with no ID and every stage refused to
  // run. That gap is the only reason a manual "Generate ID" button existed.
  // Generating on demand here means every stage self-heals: an ID exists by the
  // time any stage reads the row, no matter how the row was created.
  let id = sheet.getRange(row, COL_IDEA.ID).getValue();
  if (!id || id.toString().trim() === "") {
    const company = sheet.getRange(row, COL_IDEA.COMPANY).getValue();
    if (!company || company.toString().trim() === "") {
      SpreadsheetApp.getUi().alert(
        "This row is empty. Enter a Company name first — the ID is created automatically."
      );
      return null;
    }
    // The domain code is baked into the ID and the ID is the join key for every
    // sheet, Drive folder and tracker row — so never guess it. Resolve from the
    // Domain column or Company_Master; if neither can tell, ASK, because a wrong
    // code is far more expensive to unpick later than one dialog now.
    let resolved = resolveDomainCode_(sheet, row);
    if (!resolved) {
      const ui = SpreadsheetApp.getUi();
      const choice = ui.prompt(
        "Which domain is " + company.toString().trim() + "?",
        "The ID includes a domain code and cannot be changed easily afterwards.\n\n" +
        "  1  Business            (finance, retail, energy, industrial, health…)\n" +
        "  2  Sports\n" +
        "  3  Media & Creator Economy\n" +
        "  4  Public Sector\n" +
        "  5  Startups & Tech\n\n" +
        "Enter 1–5:",
        ui.ButtonSet.OK_CANCEL
      );
      if (choice.getSelectedButton() !== ui.Button.OK) return null;
      const pick = { "1": "Business", "2": "Sports", "3": "Media & Creator Economy",
                     "4": "Public Sector", "5": "Startups & Tech" }[choice.getResponseText().trim()];
      if (!pick) { ui.alert("No ID created — enter 1–5."); return null; }
      sheet.getRange(row, COL_IDEA.DOMAIN).setValue(pick);   // remember the answer
      resolved = { code: DOMAIN_CODES[pick], how: "you chose " + pick };
    }

    id = generateId(sheet, row, resolved.code);
    sheet.getRange(row, COL_IDEA.ID).setValue(id);
    SpreadsheetApp.flush();
    Logger.log("getActiveIdeaRow: auto-generated ID " + id + " for row " + row + " — domain from " + resolved.how);
  }

  return {
    id             : id.toString().trim(),
    company        : sheet.getRange(row, COL_IDEA.COMPANY       ).getValue(),
    domain         : sheet.getRange(row, COL_IDEA.DOMAIN        ).getValue(),
    industry       : sheet.getRange(row, COL_IDEA.INDUSTRY      ).getValue(),
    field          : sheet.getRange(row, COL_IDEA.FIELD         ).getValue(),
    type           : sheet.getRange(row, COL_IDEA.TYPE          ).getValue(),
    initialAngle   : sheet.getRange(row, COL_IDEA.INITIAL_ANGLE ).getValue(),
    priority       : sheet.getRange(row, COL_IDEA.PRIORITY      ).getValue(),
    language       : sheet.getRange(row, COL_IDEA.LANGUAGE_FLAG ).getValue(),
    targetFormat   : sheet.getRange(row, COL_IDEA.TARGET_FORMAT ).getValue() || "Standard (4–7 min)",
    callToAction   : sheet.getRange(row, COL_IDEA.CALL_TO_ACTION).getValue() || "Subscribe",
    series         : sheet.getRange(row, COL_IDEA.SERIES        ).getValue() || "",
    note           : sheet.getRange(row, COL_IDEA.NOTE          ).getValue(),
    productionMode : "Remotion Only",  // fixed — data-driven model, no AI video
    maxAiClips     : 0,                 // retired — no KlingAI scenes generated
    projectName    : ""                  // retired — CapCut replaced by Shotstack
  };
}

// ── Shared: update pipeline status column in Idea Catalogue ──────────────────
// stageKey examples: "S1", "S2", "S3", "S4", "S45", "S8D", "S10", "S11"
// statusChar: "✅" (done), "❌" (failed), "⏳" (in progress)
function updatePipelineStatus_(contentId, stageKey, statusChar) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.IDEA);
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_IDEA.ID - 1].toString().trim() !== contentId) continue;

      const cell    = sheet.getRange(i + 1, COL_IDEA.PIPELINE_STATUS);
      const current = cell.getValue().toString().trim();

      // Remove any existing token for this stage (any status char suffix)
      const tokens  = current ? current.split(" ").filter(t => t && !t.startsWith(stageKey)) : [];
      tokens.push(stageKey + (statusChar || "✅"));

      // Sort tokens by stage order
      const order = ["S1","S2","S3","S4","S45","S5","S6","S7","S8D","S9","S10","S11"];
      tokens.sort((a, b) => {
        const ak = a.replace(/[^A-Z0-9]/gi, "");
        const bk = b.replace(/[^A-Z0-9]/gi, "");
        const ai = order.indexOf(ak.replace(/[✅❌⏳]/g, ""));
        const bi = order.indexOf(bk.replace(/[✅❌⏳]/g, ""));
        return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
      });

      cell.setValue(tokens.join(" "));
      SpreadsheetApp.flush();
      return;
    }
  } catch (e) {
    Logger.log("updatePipelineStatus_: " + e.message);
  }
}

// ── Shared: check if previous stage is complete ───────────────────────────────
function checkPreviousStage(contentId, requiredSheet, stageName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(requiredSheet);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      "Sheet '" + requiredSheet + "' not found. " +
      "Please complete Stage: " + stageName + " first."
    );
    return false;
  }

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === contentId) {
      return true;
    }
  }

  SpreadsheetApp.getUi().alert(
    "No entry found for ID: " + contentId + " in " + requiredSheet + ".\n" +
    "Please complete " + stageName + " first."
  );
  return false;
}

// ── Shared: get research sources for a content ID ─────────────────────────────
function getResearchSources(contentId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.RESEARCH);
  if (!sheet) return [];

  const data         = sheet.getDataRange().getValues();
  const allRows      = [];
  const approvedRows = [];
  const rejectedRows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_RESEARCH.ID - 1].toString().trim() !== contentId) continue;

    const source = {
      row          : i + 1,
      sourceType   : data[i][COL_RESEARCH.SOURCE_TYPE    - 1],
      details      : data[i][COL_RESEARCH.DETAILS        - 1],
      sourceLink   : data[i][COL_RESEARCH.SOURCE_LINK    - 1],
      keyInsight   : data[i][COL_RESEARCH.KEY_INSIGHT    - 1],
      evidenceType : data[i][COL_RESEARCH.EVIDENCE_TYPE  - 1],
      relevance    : data[i][COL_RESEARCH.RELEVANCE      - 1],
      note         : data[i][COL_RESEARCH.NOTE           - 1],
      usedInScript : data[i][COL_RESEARCH.USED_IN_SCRIPT - 1].toString().trim().toUpperCase()
    };

    allRows.push(source);
    if (source.usedInScript === "YES") approvedRows.push(source);
    if (source.usedInScript === "NO")  rejectedRows.push(source);
  }

  // ── 3-state editorial filter ──────────────────────────────────────────────
  // YES marked   → use only approved rows (your curation, mandatory)
  // All NO       → block Stage 3, alert user — nothing to build from
  // Mix of NO+blank → pass only the blank (unrejected) rows to Claude
  // All blank    → pass everything to Claude, Claude decides

  // Case 1: some YES marked → use only approved
  if (approvedRows.length > 0) {
    Logger.log("Research: " + approvedRows.length + " manually approved sources.");
    return approvedRows;
  }

  // Case 2: everything explicitly rejected → nothing usable
  if (rejectedRows.length === allRows.length && allRows.length > 0) {
    SpreadsheetApp.getUi().alert(
      "⚠️ No Usable Sources",
      "All research sources for " + contentId + " are marked NO.\n\n" +
      "Stage 3 cannot run without at least one approved or unreviewed source.\n\n" +
      "Please mark at least one source as YES or clear its NO marking.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null; // signals calling function to abort
  }

  // Case 3: mix of NO and blank → pass only unrejected rows
  const unrejectedRows = allRows.filter(s => s.usedInScript !== "NO");
  if (unrejectedRows.length < allRows.length) {
    Logger.log("Research: " + rejectedRows.length + " sources rejected. Passing " + unrejectedRows.length + " unrejected sources.");
    return unrejectedRows;
  }

  // Case 4: all blank → pass everything, Claude decides
  Logger.log("Research: no manual selections — passing all " + allRows.length + " sources.");
  return allRows;
}

// ── Shared: get master content for a content ID ───────────────────────────────
function getMasterContent(contentId) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.MASTER);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_MASTER.ID - 1].toString().trim() === contentId) {
      return {
        title          : data[i][COL_MASTER.TITLE          - 1],
        domain         : data[i][COL_MASTER.DOMAIN         - 1],
        industry       : data[i][COL_MASTER.INDUSTRY       - 1],
        field          : data[i][COL_MASTER.FIELD          - 1],
        type           : data[i][COL_MASTER.TYPE           - 1],
        hook           : data[i][COL_MASTER.HOOK           - 1],
        reverseAngle   : data[i][COL_MASTER.REVERSE_ANGLE  - 1],
        primaryAngle   : data[i][COL_MASTER.PRIMARY_ANGLE  - 1],
        discipline     : data[i][COL_MASTER.DISCIPLINE     - 1],
        coreInsight    : data[i][COL_MASTER.CORE_INSIGHT   - 1],
        checkpoints    : data[i][COL_MASTER.CHECKPOINTS    - 1],
        targetAudience : data[i][COL_MASTER.TARGET_AUDIENCE- 1],
        language       : data[i][COL_MASTER.LANGUAGE_FLAG  - 1],
        series         : data[i][COL_MASTER.SERIES         - 1]
      };
    }
  }
  return null;
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 1 — Generate Master Content Table
// ════════════════════════════════════════════════════════════════════════════════
function generateMasterContent() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Stage 1 — Generating Master Content",
    "Claude is analyzing: " + idea.company + "\nThis may take 15–30 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildMasterContentPrompt(idea);
    const raw    = callClaude(prompt, "stage_1_master");

    const validation = validateOutput("MASTER", raw, idea.id, null);
    const proceed    = showValidationResult("MASTER", validation, idea.id);

    if (!proceed) {
      logError("Stage 1 — Master Content", idea.id, "Quality Gate Failed",
        validation.failures.join(" | "));
      return;
    }

    writeMasterContent(idea.id, raw);
    updatePipelineStatus_(idea.id, "S1", "✅");
    ui.alert(
      "✅ Stage 1 Complete",
      "Master Content Table has been filled for: " + idea.id + "\n\n" +
      "Review the Master Content Table tab, then run Stage 2 — Generate Research.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S1", "❌");
    logError("Stage 1 — Master Content", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 1 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

function buildMasterContentPrompt(idea) {
  return `
${readChannelMemory()}
You are generating the Master Content Table entry for a GovernX YouTube video.

CONTENT TO ANALYZE:
Company/Topic  : ${idea.company}
Domain         : ${idea.domain}
Industry       : ${idea.industry}
Field          : ${idea.field}
Content Type   : ${idea.type}
Initial Angle  : ${idea.initialAngle}
Language Flag  : ${idea.language}
Series         : ${idea.series || "To be determined by analysis"}
Creator Notes  : ${idea.note || "None"}

INSTRUCTION: If Creator Notes are present, treat them as directional input
that should influence the angle, audience, framing, or discipline focus.
Do not ignore them.

${idea.language === "Arabic" ? `
ARABIC MODE — MANDATORY:
Write the following fields in Modern Standard Arabic (فصحى):
TITLE, HOOK, FINAL_MOMENT, CORE_INSIGHT, CHECKPOINTS, THUMBNAIL_BRIEF
Keep in English: DOMAIN, INDUSTRY, FIELD, TYPE, DISCIPLINE, REVERSE_ANGLE,
PRIMARY_ANGLE, TARGET_AUDIENCE, SERIES
Numbers and stats stay as numerals (2013, $7.2B) — do not write in Arabic words.
` : idea.language === "Bilingual" ? `
BILINGUAL MODE — MANDATORY:
HOOK: write in Arabic first, then English on the same line separated by " | "
FINAL_MOMENT: write in Arabic only
All other fields: write in English
` : ""}

VALID TAXONOMY:
Domains    : Business | Sports | Media & Creator Economy | Public Sector | Startups & Tech
Industries : Technology | Finance & Banking | Healthcare | Automotive | Education | 
             Food & Beverage | Fashion & Apparel | Hospitality & Tourism | 
             Media & Entertainment | Sports | Retail & E-Commerce | 
             Energy & Utilities | Telecom | Government & Public Sector
Disciplines: GRC | BPR | GRC+BPR
Types      : Success | Failure | Turning Point | Collapse | Underdog
Series     : Governance Collapse Series | BPR Turning Points | Risk Blind Spots | 
             Leadership Decisions | System Design Failures

TASK:
Using the GovernX reverse-engineering method, generate the Master Content entry.
Apply all legal guardrails and quality standards from your system instructions.
${idea.series ? "IMPORTANT: Series has been pre-assigned as '" + idea.series + "' — use this value exactly." : ""}

CRITICAL FORMAT RULES:
- Every field must be on a SINGLE LINE immediately after the field name and colon
- Do NOT wrap values onto multiple lines
- Do NOT add blank lines between fields
- Return ONLY the fields listed below — no extra commentary

Return your response in EXACTLY this format (one value per line, no extra text):

TITLE: [full video title]
DOMAIN: [single value from taxonomy]
INDUSTRY: [single value from taxonomy]
FIELD: [specific sub-sector]
TYPE: [single value from taxonomy]
HOOK: [1 sentence — the final visible outcome, specific date/number/event]
FINAL_MOMENT: [same hook ${idea.language === "Arabic" || idea.language === "Bilingual" ? "in Arabic" : "in English"}]
REVERSE_ANGLE: [all possible angles, pipe-separated]
PRIMARY_ANGLE: [single chosen angle for this video]
DISCIPLINE: [GRC | BPR | GRC+BPR]
CORE_INSIGHT: [1 sentence — the GovernX lesson]
CHECKPOINTS: [reverse engineering steps, arrow-separated e.g. Collapse → Warning → Shock → Root Cause]
TARGET_AUDIENCE: [C-Suite / Executive | Board Member | Mid-Level Manager | Compliance & Risk Professional | Business Owner / Founder | Business Student / Consultant]
LANGUAGE_FLAG: ${idea.language}
SERIES: ${idea.series || "[choose most relevant series from taxonomy]"}
THUMBNAIL_BRIEF: [Specific thumbnail visual description — include: main text overlay (max 5 words, high contrast), background image concept, color treatment, emotion/mood, what makes someone stop scrolling. Format: "TEXT: [words] | IMAGE: [concept] | COLOR: [treatment] | HOOK: [why it stops scrolling]"]

${ROOT_CAUSE_ANALYST_SKILL}
`;
}

function writeMasterContent(contentId, raw) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.MASTER);

  // ── Robust parser: captures value up to next ALL_CAPS field or end ────────
  // This handles cases where Claude writes long values that span lines
  const getValue = (field) => {
    const pattern = new RegExp(field + ":\\s*([\\s\\S]*?)(?=\\n[A-Z_]{2,}:|$)");
    const match   = raw.match(pattern);
    if (!match) return "";
    // Collapse newlines and trim — single-value fields should be one line
    return match[1].replace(/\n/g, " ").trim();
  };

  // ── Log all parsed values for debugging ──────────────────────────────────
  Logger.log("=== MASTER CONTENT PARSE DEBUG ===");
  Logger.log("RAW LENGTH: " + raw.length);
  const fields = ["TITLE","DOMAIN","INDUSTRY","FIELD","TYPE","HOOK","FINAL_MOMENT",
                  "REVERSE_ANGLE","PRIMARY_ANGLE","DISCIPLINE","CORE_INSIGHT",
                  "CHECKPOINTS","TARGET_AUDIENCE","LANGUAGE_FLAG","SERIES"];
  fields.forEach(f => Logger.log(f + ": [" + getValue(f) + "]"));
  Logger.log("=== END DEBUG ===");

  // ── Upsert logic: update existing row if ID found, create new row if not ──
  let targetRow = -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_MASTER.ID - 1].toString().trim() === contentId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
  } else {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Entry Already Exists",
      "Master Content already has an entry for: " + contentId + "\n\nOverwrite it with the new generation?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // ── Clear validation on ALL Claude-written columns before writing ─────────
  const freeTextCols = [
    COL_MASTER.TITLE,
    COL_MASTER.FIELD,
    COL_MASTER.HOOK,
    COL_MASTER.FINAL_MOMENT,
    COL_MASTER.REVERSE_ANGLE,
    COL_MASTER.PRIMARY_ANGLE,
    COL_MASTER.DISCIPLINE,
    COL_MASTER.CORE_INSIGHT,
    COL_MASTER.CHECKPOINTS,
    COL_MASTER.TARGET_AUDIENCE,
    COL_MASTER.LANGUAGE_FLAG,
    COL_MASTER.SERIES,
    COL_MASTER.THUMBNAIL_BRIEF
  ];
  freeTextCols.forEach(col => {
    sheet.getRange(targetRow, col).clearDataValidations();
  });

  sheet.getRange(targetRow, COL_MASTER.ID             ).setValue(contentId);
  sheet.getRange(targetRow, COL_MASTER.TITLE          ).setValue(getValue("TITLE"));
  sheet.getRange(targetRow, COL_MASTER.DOMAIN         ).setValue(getValue("DOMAIN"));
  sheet.getRange(targetRow, COL_MASTER.INDUSTRY       ).setValue(getValue("INDUSTRY"));
  sheet.getRange(targetRow, COL_MASTER.FIELD          ).setValue(getValue("FIELD"));
  sheet.getRange(targetRow, COL_MASTER.TYPE           ).setValue(getValue("TYPE"));
  sheet.getRange(targetRow, COL_MASTER.HOOK           ).setValue(getValue("HOOK"));
  sheet.getRange(targetRow, COL_MASTER.FINAL_MOMENT   ).setValue(getValue("FINAL_MOMENT"));
  sheet.getRange(targetRow, COL_MASTER.REVERSE_ANGLE  ).setValue(getValue("REVERSE_ANGLE"));
  sheet.getRange(targetRow, COL_MASTER.PRIMARY_ANGLE  ).setValue(getValue("PRIMARY_ANGLE"));
  sheet.getRange(targetRow, COL_MASTER.DISCIPLINE     ).setValue(getValue("DISCIPLINE"));
  sheet.getRange(targetRow, COL_MASTER.CORE_INSIGHT   ).setValue(getValue("CORE_INSIGHT"));
  sheet.getRange(targetRow, COL_MASTER.CHECKPOINTS    ).setValue(getValue("CHECKPOINTS"));
  sheet.getRange(targetRow, COL_MASTER.TARGET_AUDIENCE).setValue(getValue("TARGET_AUDIENCE"));
  sheet.getRange(targetRow, COL_MASTER.LANGUAGE_FLAG  ).setValue(getValue("LANGUAGE_FLAG"));
  sheet.getRange(targetRow, COL_MASTER.SERIES         ).setValue(getValue("SERIES"));
  sheet.getRange(targetRow, COL_MASTER.THUMBNAIL_BRIEF).setValue(getValue("THUMBNAIL_BRIEF"));
  // NOTE column is never touched — always preserved as your manual input
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 2 — Generate Research Database
// ════════════════════════════════════════════════════════════════════════════════
function generateResearchDatabase() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.MASTER, "Stage 1 — Master Content")) return;

  const master = getMasterContent(idea.id);
  const ui     = SpreadsheetApp.getUi();

  ui.alert(
    "Stage 2 — Generating Research Database",
    "Claude is identifying sources for: " + idea.company + "\nThis may take 20–40 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildResearchPrompt(idea, master);
    const raw    = callClaude(prompt, "stage_2_research");

    const validation = validateOutput("RESEARCH", raw, idea.id, null);
    const proceed    = showValidationResult("RESEARCH", validation, idea.id);

    if (!proceed) {
      logError("Stage 2 — Research", idea.id, "Quality Gate Failed",
        validation.failures.join(" | "));
      return;
    }

    writeResearchDatabase(idea.id, idea.company, raw);
    updatePipelineStatus_(idea.id, "S2", "✅");
    ui.alert(
      "✅ Stage 2 Complete",
      "Research Database has been filled for: " + idea.id + "\n\n" +
      "Review the Research Database tab.\n" +
      "Add source links where marked 'Search: [query]'.\n\n" +
      "NEXT — verify the research before scripting (GovernX menu):\n" +
      "  🔎 Research ② Run verified research…\n" +
      "  🔎 Research ③ Approve ticked claims → Data Moments\n" +
      "  🔎 Research ⑤ Publish → Research Database (feeds Stage 3)\n\n" +
      "Then run Stage 3 — Generate Script.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S2", "❌");
    logError("Stage 2 — Research", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 2 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

function buildResearchPrompt(idea, master) {
  return `
You are building the Research Database for a GovernX YouTube video.

CONTENT BRIEF:
Company/Topic  : ${idea.company}
Title          : ${master.title}
Primary Angle  : ${master.primaryAngle}
Discipline     : ${master.discipline}
Core Insight   : ${master.coreInsight}
Checkpoints    : ${master.checkpoints}

TASK:
Identify 6–8 high-quality sources that provide evidence for the GovernX 
reverse-engineering analysis of this story. Sources must directly support 
the Primary Angle and Discipline identified above.

Apply all legal guardrails: only suggest real, verifiable sources. 
If you cannot confirm a source URL, write "Search: [recommended search query]" 
in the Source Link field.

VALID VALUES:
Source Types   : Book | Article | Interview | Report | Video | Academic Reports
Evidence Types : Analysis | Quote | Data | Case Study | Research | Insider | Primary | Opinion
Relevance      : High | Medium | Low

Return each source in EXACTLY this format:

SOURCE_1_START
SOURCE_TYPE: [value]
DETAILS: [Title — Author/Publisher/Platform]
SOURCE_LINK: [URL or "Search: [query]"]
KEY_INSIGHT: [1 sentence — what this source proves for the GovernX angle]
EVIDENCE_TYPE: [value]
RELEVANCE: [High | Medium | Low]
SOURCE_1_END

SOURCE_2_START
[repeat for each source]
SOURCE_N_END

${DATA_RESEARCH_SKILL}
`;
}

function writeResearchDatabase(contentId, topic, raw) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.RESEARCH);

  // ── Upsert: delete any existing rows for this contentId first ────────────
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][COL_RESEARCH.ID - 1].toString().trim() === contentId) {
      rowsToDelete.push(i + 1);
    }
  }
  if (rowsToDelete.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Research Already Exists",
      "Found " + rowsToDelete.length + " existing research rows for: " + contentId + "\n\nOverwrite with new generation?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    rowsToDelete.forEach(row => sheet.deleteRow(row));
  }

  // Parse all SOURCE blocks and write fresh
  const sourceBlocks = raw.match(/SOURCE_\d+_START([\s\S]*?)SOURCE_\d+_END/g) || [];

  sourceBlocks.forEach(block => {
    const getValue = (field) => {
      const match = block.match(new RegExp(field + ":\\s*(.+)"));
      return match ? match[1].trim() : "";
    };

    const lastRow = sheet.getLastRow() + 1;

    // Clear validation on free-text fields before writing
    [COL_RESEARCH.DETAILS, COL_RESEARCH.SOURCE_LINK, COL_RESEARCH.KEY_INSIGHT].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });

    sheet.getRange(lastRow, COL_RESEARCH.ID            ).setValue(contentId);
    sheet.getRange(lastRow, COL_RESEARCH.TOPIC         ).setValue(topic);
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_TYPE   ).setValue(getValue("SOURCE_TYPE"));
    sheet.getRange(lastRow, COL_RESEARCH.DETAILS       ).setValue(getValue("DETAILS"));
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_LINK   ).setValue(getValue("SOURCE_LINK"));
    sheet.getRange(lastRow, COL_RESEARCH.KEY_INSIGHT   ).setValue(getValue("KEY_INSIGHT"));
    sheet.getRange(lastRow, COL_RESEARCH.EVIDENCE_TYPE ).setValue(getValue("EVIDENCE_TYPE"));
    sheet.getRange(lastRow, COL_RESEARCH.TIMESTAMP     ).setValue(new Date());
    sheet.getRange(lastRow, COL_RESEARCH.RELEVANCE     ).setValue(getValue("RELEVANCE"));
    sheet.getRange(lastRow, COL_RESEARCH.USED_IN_SCRIPT).setValue("");
  });

  // ── Parse DATA blocks and write as Data-type rows ─────────────────────────
  // DATA rows use SOURCE_TYPE = "Data" and store the structured payload in NOTE
  const dataBlocks = raw.match(/DATA_\d+_START([\s\S]*?)DATA_\d+_END/g) || [];

  dataBlocks.forEach(block => {
    const getValue = (field) => {
      const match = block.match(new RegExp(field + ":\\s*(.+)"));
      return match ? match[1].trim() : "";
    };

    const label          = getValue("DATA_LABEL");
    const value          = getValue("DATA_VALUE");
    const year           = getValue("DATA_YEAR");
    const context        = getValue("DATA_CONTEXT");
    const compareLabel   = getValue("COMPARE_LABEL");
    const compareValue   = getValue("COMPARE_VALUE");
    const compareYear    = getValue("COMPARE_YEAR");
    const compareContext = getValue("COMPARE_CONTEXT");
    const vizType        = getValue("VIZ_TYPE");
    const sourceLink     = getValue("SOURCE_LINK");
    const keyInsight     = getValue("KEY_INSIGHT");

    if (!label || !value) return; // skip malformed blocks

    // Build machine-readable NOTE payload for Stage 4 to consume
    const notePayload = [
      "DATA_LABEL:"     + label,
      "DATA_VALUE:"     + value,
      "DATA_YEAR:"      + year,
      "DATA_CONTEXT:"   + context,
      "COMPARE_LABEL:"  + compareLabel,
      "COMPARE_VALUE:"  + compareValue,
      "COMPARE_YEAR:"   + compareYear,
      "COMPARE_CONTEXT:"+ compareContext,
      "VIZ_TYPE:"       + vizType
    ].join(" | ");

    const lastRow = sheet.getLastRow() + 1;

    [COL_RESEARCH.DETAILS, COL_RESEARCH.SOURCE_LINK,
     COL_RESEARCH.KEY_INSIGHT, COL_RESEARCH.NOTE].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });

    sheet.getRange(lastRow, COL_RESEARCH.ID            ).setValue(contentId);
    sheet.getRange(lastRow, COL_RESEARCH.TOPIC         ).setValue(topic);
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_TYPE   ).setValue("Data");
    sheet.getRange(lastRow, COL_RESEARCH.DETAILS       ).setValue(label + " (" + year + (compareYear && compareYear !== "N/A" ? " → " + compareYear : "") + ")");
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_LINK   ).setValue(sourceLink);
    sheet.getRange(lastRow, COL_RESEARCH.KEY_INSIGHT   ).setValue(keyInsight);
    sheet.getRange(lastRow, COL_RESEARCH.EVIDENCE_TYPE ).setValue("Data");
    sheet.getRange(lastRow, COL_RESEARCH.TIMESTAMP     ).setValue(new Date());
    sheet.getRange(lastRow, COL_RESEARCH.RELEVANCE     ).setValue("High");
    sheet.getRange(lastRow, COL_RESEARCH.USED_IN_SCRIPT).setValue("");
    sheet.getRange(lastRow, COL_RESEARCH.NOTE          ).setValue(notePayload);

    // Visual distinction — light blue background for Data rows
    sheet.getRange(lastRow, 1, 1, 11).setBackground("#E3F2FD");
  });

  // ── Parse RISK blocks ─────────────────────────────────────────────────────
  const riskBlocks = raw.match(/RISK_\d+_START([\s\S]*?)RISK_\d+_END/g) || [];

  riskBlocks.forEach(block => {
    const get = (field) => {
      const match = block.match(new RegExp(field + ":\\s*(.+)"));
      return match ? match[1].trim() : "";
    };

    const label       = get("RISK_LABEL");
    const likelihood  = get("RISK_LIKELIHOOD");
    const impact      = get("RISK_IMPACT");
    const highlight   = get("RISK_HIGHLIGHT");
    const description = get("RISK_DESCRIPTION");
    const sourceLink  = get("SOURCE_LINK");

    if (!label || !likelihood || !impact) return;

    const notePayload = [
      "RISK_LABEL:"       + label,
      "RISK_LIKELIHOOD:"  + likelihood,
      "RISK_IMPACT:"      + impact,
      "RISK_HIGHLIGHT:"   + highlight,
      "RISK_DESCRIPTION:" + description
    ].join(" | ");

    const lastRow = sheet.getLastRow() + 1;
    [COL_RESEARCH.DETAILS, COL_RESEARCH.SOURCE_LINK,
     COL_RESEARCH.KEY_INSIGHT, COL_RESEARCH.NOTE].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });
    sheet.getRange(lastRow, COL_RESEARCH.ID            ).setValue(contentId);
    sheet.getRange(lastRow, COL_RESEARCH.TOPIC         ).setValue(topic);
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_TYPE   ).setValue("Risk");
    sheet.getRange(lastRow, COL_RESEARCH.DETAILS       ).setValue(label + " (L" + likelihood + "/I" + impact + ")");
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_LINK   ).setValue(sourceLink);
    sheet.getRange(lastRow, COL_RESEARCH.KEY_INSIGHT   ).setValue(description);
    sheet.getRange(lastRow, COL_RESEARCH.EVIDENCE_TYPE ).setValue("Risk");
    sheet.getRange(lastRow, COL_RESEARCH.TIMESTAMP     ).setValue(new Date());
    sheet.getRange(lastRow, COL_RESEARCH.RELEVANCE     ).setValue("High");
    sheet.getRange(lastRow, COL_RESEARCH.USED_IN_SCRIPT).setValue("");
    sheet.getRange(lastRow, COL_RESEARCH.NOTE          ).setValue(notePayload);
    sheet.getRange(lastRow, 1, 1, 11).setBackground("#FCE4EC");
  });

  // ── Parse KPI blocks ──────────────────────────────────────────────────────
  const kpiBlocks = raw.match(/KPI_\d+_START([\s\S]*?)KPI_\d+_END/g) || [];

  kpiBlocks.forEach(block => {
    const get = (field) => {
      const match = block.match(new RegExp(field + ":\\s*(.+)"));
      return match ? match[1].trim() : "";
    };

    const label     = get("KPI_LABEL");
    const value     = get("KPI_VALUE");
    const trend     = get("KPI_TREND");
    const change    = get("KPI_CHANGE");
    const context   = get("KPI_CONTEXT");
    const highlight = get("KPI_HIGHLIGHT");
    const sourceLink= get("SOURCE_LINK");

    if (!label || !value) return;

    const notePayload = [
      "KPI_LABEL:"    + label,
      "KPI_VALUE:"    + value,
      "KPI_TREND:"    + trend,
      "KPI_CHANGE:"   + change,
      "KPI_CONTEXT:"  + context,
      "KPI_HIGHLIGHT:"+ highlight
    ].join(" | ");

    const lastRow = sheet.getLastRow() + 1;
    [COL_RESEARCH.DETAILS, COL_RESEARCH.SOURCE_LINK,
     COL_RESEARCH.KEY_INSIGHT, COL_RESEARCH.NOTE].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });
    sheet.getRange(lastRow, COL_RESEARCH.ID            ).setValue(contentId);
    sheet.getRange(lastRow, COL_RESEARCH.TOPIC         ).setValue(topic);
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_TYPE   ).setValue("KPI");
    sheet.getRange(lastRow, COL_RESEARCH.DETAILS       ).setValue(label + ": " + value + (change && change !== "N/A" ? " (" + change + ")" : ""));
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_LINK   ).setValue(sourceLink);
    sheet.getRange(lastRow, COL_RESEARCH.KEY_INSIGHT   ).setValue(context);
    sheet.getRange(lastRow, COL_RESEARCH.EVIDENCE_TYPE ).setValue("KPI");
    sheet.getRange(lastRow, COL_RESEARCH.TIMESTAMP     ).setValue(new Date());
    sheet.getRange(lastRow, COL_RESEARCH.RELEVANCE     ).setValue("High");
    sheet.getRange(lastRow, COL_RESEARCH.USED_IN_SCRIPT).setValue("");
    sheet.getRange(lastRow, COL_RESEARCH.NOTE          ).setValue(notePayload);
    sheet.getRange(lastRow, 1, 1, 11).setBackground("#E8F5E9");
  });

  // ── Parse GAUGE blocks ────────────────────────────────────────────────────
  const gaugeBlocks = raw.match(/GAUGE_\d+_START([\s\S]*?)GAUGE_\d+_END/g) || [];

  gaugeBlocks.forEach(block => {
    const get = (field) => {
      const match = block.match(new RegExp(field + ":\\s*(.+)"));
      return match ? match[1].trim() : "";
    };

    const label      = get("GAUGE_LABEL");
    const value      = get("GAUGE_VALUE");
    const unit       = get("GAUGE_UNIT");
    const context    = get("GAUGE_CONTEXT");
    const highlight  = get("GAUGE_HIGHLIGHT");
    const threshold  = get("GAUGE_THRESHOLD");
    const sourceLink = get("SOURCE_LINK");

    if (!label || !value) return;

    const notePayload = [
      "GAUGE_LABEL:"     + label,
      "GAUGE_VALUE:"     + value,
      "GAUGE_UNIT:"      + unit,
      "GAUGE_CONTEXT:"   + context,
      "GAUGE_HIGHLIGHT:" + highlight,
      "GAUGE_THRESHOLD:" + threshold
    ].join(" | ");

    const lastRow = sheet.getLastRow() + 1;
    [COL_RESEARCH.DETAILS, COL_RESEARCH.SOURCE_LINK,
     COL_RESEARCH.KEY_INSIGHT, COL_RESEARCH.NOTE].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });
    sheet.getRange(lastRow, COL_RESEARCH.ID            ).setValue(contentId);
    sheet.getRange(lastRow, COL_RESEARCH.TOPIC         ).setValue(topic);
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_TYPE   ).setValue("Gauge");
    sheet.getRange(lastRow, COL_RESEARCH.DETAILS       ).setValue(label + ": " + value + (unit || "%") + (threshold && threshold !== "N/A" ? " (threshold " + threshold + ")" : ""));
    sheet.getRange(lastRow, COL_RESEARCH.SOURCE_LINK   ).setValue(sourceLink);
    sheet.getRange(lastRow, COL_RESEARCH.KEY_INSIGHT   ).setValue(context);
    sheet.getRange(lastRow, COL_RESEARCH.EVIDENCE_TYPE ).setValue("Gauge");
    sheet.getRange(lastRow, COL_RESEARCH.TIMESTAMP     ).setValue(new Date());
    sheet.getRange(lastRow, COL_RESEARCH.RELEVANCE     ).setValue("High");
    sheet.getRange(lastRow, COL_RESEARCH.USED_IN_SCRIPT).setValue("");
    sheet.getRange(lastRow, COL_RESEARCH.NOTE          ).setValue(notePayload);
    sheet.getRange(lastRow, 1, 1, 11).setBackground("#F3E5F5");
  });

  Logger.log("Research Database written: " + sourceBlocks.length + " narrative sources + " +
    dataBlocks.length + " data + " + riskBlocks.length + " risk + " +
    kpiBlocks.length + " KPI + " + gaugeBlocks.length + " gauge rows.");
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 3 — Generate Script Bank
// ════════════════════════════════════════════════════════════════════════════════
function generateScript() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.MASTER,   "Stage 1 — Master Content")) return;
  if (!checkPreviousStage(idea.id, SHEET.RESEARCH, "Stage 2 — Research Database")) return;

  const master  = getMasterContent(idea.id);
  const sources = getResearchSources(idea.id);
  if (sources === null) return; // all sources rejected — user was already alerted
  const ui      = SpreadsheetApp.getUi();

  // Read Target Format directly from Idea Catalogue — not guessed from content type
  const targetFormat = idea.targetFormat || "Standard (4–7 min)";

  ui.alert(
    "Stage 3 — Generating Script",
    "Claude is writing the full GovernX script for: " + idea.company +
    "\nFormat: " + targetFormat +
    "\nThis may take 30–60 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildScriptPrompt(idea, master, sources, targetFormat);
    const raw    = callClaude(prompt, "stage_3_script");

    // ── Guard: ensure Claude returned a non-empty script ─────────────────────
    if (!raw || raw.trim().length < 100) {
      throw new Error(
        "Claude returned an empty or near-empty response for Stage 3.\n" +
        "Check your Anthropic API credits and retry."
      );
    }

    // ── QA Critic evaluation (Skill 5) ────────────────────────────────────────
    // A second Claude instance (QA_CRITIC_SYSTEM_CONTEXT) evaluates the script
    // independently before it is written to the sheet.
    // This replaces the old validateOutput/showValidationResult block for scripts.
    const evalRaw    = callClaudeAsEvaluator(raw, idea, master, sources);
    const evalResult = parseEvalReport(evalRaw);
    const proceed    = showEvalResult(evalResult, idea);

    if (!proceed) {
      logError("Stage 3 — Script", idea.id,
        "QA Critic: " + evalResult.verdict, evalResult.summary);
      return;
    }

    // ── Save QA scores to Script Bank NOTE column for Channel Memory ──────────
    saveQAScoresToScript(idea.id, evalResult);

    writeScript(idea.id, raw, targetFormat, master, idea);
    // Source marking: use sources_used from QC block if critic approved
    const qcMatch = raw.match(/sources_used:\s*(.+)/);
    if (qcMatch) markSourcesUsed(idea.id, qcMatch[1].trim());

    updatePipelineStatus_(idea.id, "S3", "✅");
    ui.alert(
      "✅ Stage 3 Complete",
      "Script Bank has been filled for: " + idea.id + "\n\n" +
      "Used sources have been marked in Research Database.\n" +
      "Review the Script Bank tab, then run Stage 4 — Generate Scenes.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S3", "❌");
    logError("Stage 3 — Script", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 3 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 3B — QA Review (runs separately to avoid 6-min timeout)
// Reads the saved script from Script Bank and runs the QA Critic on it
// ════════════════════════════════════════════════════════════════════════════════
function reviewScript() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;

  const master  = getMasterContent(idea.id);
  const sources = getResearchSources(idea.id);
  const ui      = SpreadsheetApp.getUi();

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let   savedScript = "";

  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      savedScript =
        "VOICEOVER_SCRIPT:\n" + scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1] + "\n\n" +
        "GRC_BPR_CLOSING:\n" + scriptData[i][COL_SCRIPT.GRC_BPR_CLOSING   - 1] + "\n\n" +
        "HOOK: " + scriptData[i][COL_SCRIPT.HOOK - 1] + "\n\n" +
        "NARRATIVE_FLOW:\n" + scriptData[i][COL_SCRIPT.NARRATIVE_FLOW - 1] + "\n\n" +
        "TARGET_FORMAT: " + scriptData[i][COL_SCRIPT.TARGET_FORMAT - 1] + "\n\n" +
        // Include NOTE (DATA_MOMENTS + SCENE_BLUEPRINT)
        scriptData[i][COL_SCRIPT.NOTE - 1].toString();
      break;
    }
  }

  if (!savedScript) {
    ui.alert("No script found for: " + idea.id + "\nRun Stage 3 first.");
    return;
  }

  ui.alert(
    "Stage 3B — QA Review",
    "Running QA Critic on saved script for: " + idea.company +
    "\n\nThis takes 30–60 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const evalRaw    = callClaudeAsEvaluator(savedScript, idea, master, sources);
    const evalResult = parseEvalReport(evalRaw);
    const proceed    = showEvalResult(evalResult, idea);

    if (!proceed) {
      logError("Stage 3B — QA Review", idea.id,
        "QA Critic: " + evalResult.verdict, evalResult.summary);
      ui.alert(
        "Script needs revision",
        "The QA Critic rejected the script.\n\n" +
        "Review the KEY FIXES in the evaluation report, then re-run Stage 3 to regenerate.",
        ui.ButtonSet.OK
      );
      return;
    }

    // Save QA scores to NOTE column for Channel Memory
    saveQAScoresToScript(idea.id, evalResult);

    // Update status to QA Approved
    for (let i = 1; i < scriptData.length; i++) {
      if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
        scriptSheet.getRange(i + 1, COL_SCRIPT.STATUS).setValue("QA Approved");
        break;
      }
    }

    ui.alert(
      "✅ Stage 3B Complete — QA Approved",
      "Script approved for: " + idea.id + "\n\n" +
      "Score: " + evalResult.totalScore + "/70  Verdict: " + evalResult.verdict + "\n\n" +
      "Proceed to Stage 4 — Generate Scenes.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 3B — QA Review", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 3B Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}


function buildScriptPrompt(idea, master, sources, targetFormat) {

  // Build source brief — label as curated if manually selected
  const isCurated = sources.some(s => s.usedInScript === "YES");
  const curationNote = isCurated
    ? "NOTE: These sources have been MANUALLY APPROVED by the creator. You MUST use ALL of them in the script. Do not substitute or skip any."
    : "NOTE: No manual curation applied. Use a minimum of 4 sources from this list.";

  const sourceBrief = sources.map((s, i) =>
    `Source ${i + 1}: [${s.sourceType}] ${s.details}\n` +
    `  Key Insight: ${s.keyInsight}\n` +
    `  Evidence Type: ${s.evidenceType} | Relevance: ${s.relevance}` +
    (s.sourceType === "Data" && s.note
      ? `\n  DATA PAYLOAD: ${s.note}`
      : "")
  ).join("\n\n");

  const wordRange = WORD_COUNT[targetFormat] || WORD_COUNT["Standard (4–7 min)"];

  // Scene count scales with duration. A 5-min video needs ~16 scenes (~20s each),
  // not the flat "minimum 4" the skill defaults to — otherwise cards hold ~40s
  // and the video drags. DATA_MOMENTS drive the data-scene count downstream, so
  // we set a format-scaled target here.
  const SCENE_TARGET = {
    "Short (< 90s)"         : { dmMin: 3,  dmMax: 4,  scenes: "~7"  },
    "Standard (4–7 min)"    : { dmMin: 10, dmMax: 12, scenes: "~16" },
    "Deep Dive (10–15 min)" : { dmMin: 18, dmMax: 24, scenes: "~28" }
  };
  const sceneTarget = SCENE_TARGET[targetFormat] || SCENE_TARGET["Standard (4–7 min)"];

  return `
You are writing the full GovernX script for a YouTube video.

CONTENT BRIEF:
Company/Topic  : ${idea.company}
Title          : ${master.title}
Language       : ${idea.language}
Target Format  : ${targetFormat} (${wordRange.min}–${wordRange.max} words)
Scene target   : ${sceneTarget.scenes} scenes → produce ${sceneTarget.dmMin}–${sceneTarget.dmMax} DATA_MOMENTS

DATA_MOMENTS REQUIREMENT (overrides the "minimum 4" default):
- Produce ${sceneTarget.dmMin}–${sceneTarget.dmMax} DATA_MOMENTS for this ${targetFormat} video. This drives the
  scene count — too few and each scene holds ~40s and the video drags.
- ONE dedicated data scene per DATA_MOMENT. Do NOT fold a figure into a
  Checkpoint card's text; if a number deserves screen time, it gets its own DM_.
- DECOMPOSE dense number paragraphs. A sentence listing many figures (penalties,
  refunds, settlement tranches) must become SEVERAL DATA_MOMENTS — one per
  figure or logical group — not one crowded scene. This also keeps the voiceover
  from reading a wall of numbers.
- USE VARIED DATA_TYPES across the moments — do not default to counters. Reach for
  BEFORE_AFTER (a figure that changed), COMPARISON / split (two sides, or a peer
  contrast), BAR_CHART (several values ranked), LINE_GRAPH (a trend over years),
  SINGLE_STAT (one dramatic number). Match the type to the shape of the data.
- Preserve the source's hedges ("potentially", "approximately") and attribution
  in every DATA_MOMENT label, exactly as in the verified evidence.
Primary Angle  : ${master.primaryAngle}
Discipline     : ${master.discipline}
Core Insight   : ${master.coreInsight}
Checkpoints    : ${master.checkpoints}
Target Audience: ${master.targetAudience}
Series         : ${master.series}
Hook           : ${master.hook}

RESEARCH DATABASE — SOURCES TO USE IN SCRIPT:
${curationNote}

${sourceBrief}

TASK:
Write the complete GovernX video script following the reverse-engineering method.
Apply all legal guardrails and quality standards from your system instructions.

${idea.language === "Arabic" ? `
ARABIC MODE — MANDATORY:
- VOICEOVER_SCRIPT: write entirely in Modern Standard Arabic (فصحى)
- GRC_BPR_CLOSING: write in Arabic
- HOOK field: write in Arabic
- ANALYSIS_FRAMEWORK and NARRATIVE_FLOW: write in English (structural/internal use)
- Numbers and stats stay as numerals (2013, $7.2B)
- Tone: sharp, authoritative Arabic — not academic or bureaucratic
- Register: suitable for C-Suite Arab executives and business professionals
` : idea.language === "Bilingual" ? `
BILINGUAL MODE — MANDATORY:
- VOICEOVER_SCRIPT: write entirely in Arabic (فصحى)
- GRC_BPR_CLOSING: write in Arabic
- HOOK: write in Arabic first, then English translation below separated by " | "
- ANALYSIS_FRAMEWORK and NARRATIVE_FLOW: write in English
` : ""}

SCRIPT STRUCTURE (mandatory):
1. INTRO — Open with the hook (specific date/number/event). Create urgency immediately.
2. BREAKDOWN — Walk through the checkpoints in reverse order (outcome first, root cause last)
3. INSIGHT — Name the governance or process failure explicitly
4. CONCLUSION — Land the GRC/BPR closing argument. Tell the audience what proper 
   governance or BPR would have changed.

Return your response in EXACTLY this format:

ANALYSIS_FRAMEWORK:
A) What happened:
[bullet points]

B) Why it happened:
[bullet points]

C) Reverse Structure:
[numbered checkpoints]

HOOK: [1–2 sentence opening — must reference specific date, number, or event]

NARRATIVE_FLOW:
[Checkpoint by checkpoint story arc]

VOICEOVER_SCRIPT:
[Full script — every word to be spoken. Written in ${idea.language}. 
${wordRange.min}–${wordRange.max} words. Narrative prose, not bullet points.]

GRC_BPR_CLOSING:
[2–3 sentences. Explicitly names the discipline. The GovernX signature ending.]

SECTIONS: Intro, Breakdown, Insight, Conclusion
TARGET_FORMAT: ${targetFormat}
CALL_TO_ACTION: ${idea.callToAction}
SERIES: ${idea.series || master.series}
VERSION: v1
STATUS: Draft

QUALITY_CHECK_START
sources_referenced: [number]
reverse_engineering_structure: YES/NO
grc_bpr_closing_argument: YES/NO
hook_opens_with_specific_fact: YES/NO
word_count: [number]
living_persons_flagged: YES/NO
sources_used: [comma-separated list of source DETAILS values used]
copyright_risks_flagged: YES/NO
QUALITY_CHECK_END

${ROOT_CAUSE_ANALYST_SKILL}

${NARRATIVE_ARCHITECT_SKILL}

${idea.language === "Arabic" || idea.language === "Bilingual" ? ARABIC_VOICE_SKILL : ""}

${DATA_SCRIPT_SKILL}
`;
}

function writeScript(contentId, raw, targetFormat, master, idea) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.SCRIPT);

  const getBlock = (field) => {
    const match = raw.match(new RegExp(field + ":\\s*([\\s\\S]*?)(?=\\n[A-Z_]+:|QUALITY_CHECK_START|$)"));
    return match ? match[1].trim() : "";
  };
  const getValue = (field) => {
    const match = raw.match(new RegExp(field + ":\\s*(.+)"));
    return match ? match[1].trim() : "";
  };

  let voiceover = getBlock("VOICEOVER_SCRIPT").replace(
    /QUALITY_CHECK_START[\s\S]*?QUALITY_CHECK_END/, ""
  ).trim();

  // Append the Call to Action chosen in Idea Catalogue col K (idea.callToAction)
  // to the END of the spoken script, so Stage 7 narrates it in the audio.
  voiceover = appendCallToAction_(voiceover, idea.callToAction, idea.language);

  // ── Upsert: find existing row or create new ───────────────────────────────
  let targetRow = -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_SCRIPT.ID - 1].toString().trim() === contentId) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
  } else {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Script Already Exists",
      "A script already exists for: " + contentId + "\n\nOverwrite with new generation?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
  }

  // Clear validation on free-text fields before writing
  [
    COL_SCRIPT.ANALYSIS_FRAMEWORK,
    COL_SCRIPT.HOOK,
    COL_SCRIPT.NARRATIVE_FLOW,
    COL_SCRIPT.VOICEOVER_SCRIPT,
    COL_SCRIPT.GRC_BPR_CLOSING
  ].forEach(col => {
    sheet.getRange(targetRow, col).clearDataValidations();
  });

  // ── Estimate duration from word count ────────────────────────────────────
  const wcMatch      = raw.match(/word_count:\s*(\d+)/);
  const wordCount    = wcMatch ? parseInt(wcMatch[1]) : 0;
  const durationSecs = wordCount > 0 ? Math.round(wordCount / 140 * 60) : 0;
  const durationStr  = durationSecs > 0
    ? Math.floor(durationSecs / 60) + ":" + String(durationSecs % 60).padStart(2, "0")
    : "";

  // ── Parse DATA_MOMENTS and SCENE_BLUEPRINT — append to NOTE (not overwrite) ──
  // NOTE col may already have QA_SCORES from saveQAScoresToScript().
  // We merge DATA_MOMENTS and SCENE_BLUEPRINT so all three coexist.
  const dmRaw   = raw.match(/DATA_MOMENTS_START([\s\S]*?)DATA_MOMENTS_END/);
  const dmBlock = dmRaw ? dmRaw[1].trim() : "";
  const sbRaw   = raw.match(/SCENE_BLUEPRINT_START([\s\S]*?)SCENE_BLUEPRINT_END/);
  const sbBlock = sbRaw ? sbRaw[1].trim() : "";

  if (dmBlock || sbBlock) {
    const existingNote = sheet.getRange(targetRow, COL_SCRIPT.NOTE).getValue().toString();
    let updatedNote    = existingNote;

    const dmTag = dmBlock ? "DATA_MOMENTS:\n" + dmBlock : "";
    if (dmTag) {
      updatedNote = updatedNote.includes("DATA_MOMENTS:")
        ? updatedNote.replace(/DATA_MOMENTS:[\s\S]*?(?=\n\nSCENE_BLUEPRINT:|\n\nQA_SCORES:|$)/, dmTag)
        : (updatedNote ? updatedNote + "\n\n" + dmTag : dmTag);
    }

    const sbTag = sbBlock ? "SCENE_BLUEPRINT:\n" + sbBlock : "";
    if (sbTag) {
      updatedNote = updatedNote.includes("SCENE_BLUEPRINT:")
        ? updatedNote.replace(/SCENE_BLUEPRINT:[\s\S]*?(?=\n\nQA_SCORES:|$)/, sbTag)
        : (updatedNote ? updatedNote + "\n\n" + sbTag : sbTag);
    }

    sheet.getRange(targetRow, COL_SCRIPT.NOTE).setValue(updatedNote).setWrap(true);
    Logger.log("writeScript: DATA_MOMENTS=" + (dmBlock ? "yes" : "no") +
               " SCENE_BLUEPRINT=" + (sbBlock ? "yes (" + (sbBlock.match(/SCENE_BP_/g)||[]).length + " scenes)" : "no"));
  }

  sheet.getRange(targetRow, COL_SCRIPT.ID                ).setValue(contentId);
  sheet.getRange(targetRow, COL_SCRIPT.ANALYSIS_FRAMEWORK).setValue(getBlock("ANALYSIS_FRAMEWORK"));
  sheet.getRange(targetRow, COL_SCRIPT.HOOK              ).setValue(getValue("HOOK"));
  sheet.getRange(targetRow, COL_SCRIPT.NARRATIVE_FLOW    ).setValue(getBlock("NARRATIVE_FLOW"));
  sheet.getRange(targetRow, COL_SCRIPT.VOICEOVER_SCRIPT  ).setValue(voiceover);
  sheet.getRange(targetRow, COL_SCRIPT.GRC_BPR_CLOSING   ).setValue(getBlock("GRC_BPR_CLOSING"));
  sheet.getRange(targetRow, COL_SCRIPT.SECTIONS          ).setValue(getValue("SECTIONS"));
  sheet.getRange(targetRow, COL_SCRIPT.TARGET_FORMAT     ).setValue(targetFormat);
  sheet.getRange(targetRow, COL_SCRIPT.DURATION          ).setValue(durationStr);
  sheet.getRange(targetRow, COL_SCRIPT.CALL_TO_ACTION    ).setValue(idea.callToAction);
  sheet.getRange(targetRow, COL_SCRIPT.VERSION           ).setValue("v1");
  sheet.getRange(targetRow, COL_SCRIPT.STATUS            ).setValue("Draft");
  sheet.getRange(targetRow, COL_SCRIPT.SERIES            ).setValue(idea.series || master.series);
}

// ── Append the chosen Call to Action to the end of the voiceover script ───────
// Reads the value the user picked in Idea Catalogue col K (passed in as
// idea.callToAction). Appended verbatim as the final spoken line so it is
// narrated in the Stage 7 audio and shown in the Script Bank. Language-safe
// (Arabic/English punctuation) and idempotent (won't double-append on re-runs).
function appendCallToAction_(voiceover, cta, language) {
  const base  = (voiceover || "").toString().trim();
  const clean = (cta || "").toString().trim();
  if (!clean) return base;

  // Ensure the CTA reads as a complete sentence for natural TTS pacing.
  const ctaSentence = /[.!?…؟]$/.test(clean) ? clean : clean + ".";

  // Don't add it twice if a previous generation already appended this CTA.
  if (base.endsWith(ctaSentence)) return base;

  return base ? base + "\n\n" + ctaSentence : ctaSentence;
}

function markSourcesUsed(contentId, sourcesUsedStr) {
  if (!sourcesUsedStr) return;

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.RESEARCH);
  const data  = sheet.getDataRange().getValues();

  const usedSources = sourcesUsedStr.split(",").map(s => s.trim().toLowerCase());

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_RESEARCH.ID - 1].toString().trim() !== contentId) continue;
    const details = data[i][COL_RESEARCH.DETAILS - 1].toString().toLowerCase();
    const isUsed  = usedSources.some(u => details.includes(u) || u.includes(details.substring(0, 20)));
    if (isUsed) {
      sheet.getRange(i + 1, COL_RESEARCH.USED_IN_SCRIPT).setValue("YES");
    }
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 4 — Generate Visual Library
// NOTE: generateScenes() lives in Stage4_upgraded.gs (two-pass Director-embedded
// version). This file only defines the shared buildScenesPrompt() helper used by
// both Stage4_upgraded.gs and the legacy scene prompt builder.
// ════════════════════════════════════════════════════════════════════════════════

/* ── Evidence guardrail ───────────────────────────────────────────────────────
   The Master Content brief and checkpoint list are written at Stage 1, BEFORE
   verification, so they can contain figures the research engine never confirmed
   (e.g. a blended "$185M penalty", a later "$3B DOJ/SEC settlement"). Stage 4
   builds checkpoint cards from that brief and was leaking those numbers onto
   screen — contradicting the verified voiceover.

   This block injects the ONLY figures cleared for the video (the Research
   Database rows marked USED_IN_SCRIPT = YES) plus hard rules, so the director
   cannot place an unverified number, drop a hedge, or misattribute a figure.
   Injected into BOTH director passes. If a video has no verified research, the
   block is empty and behaviour is unchanged (backward-compatible).
   ──────────────────────────────────────────────────────────────────────────── */
function buildVerifiedFiguresBlock_(contentId) {
  var rows;
  try { rows = getResearchSources(contentId) || []; } catch (e) { rows = []; }
  if (!rows.length) return "";

  // Compact whitelist: one short figure per claim, grouped by who asserts it.
  // Full quotes are NOT included — the model already has the voiceover for
  // context, and a bulky block disrupts the structured pass-2 output.
  var groups = {};   // attribution → [short figures]
  rows.forEach(function (r) {
    var attr = (r.sourceType || "Unclear").toString().trim() || "Unclear";
    var body = (r.keyInsight || r.details || "").toString().trim();
    // The published KEY_INSIGHT ends with "  [value]"; prefer that, else a short slice.
    var mb = body.match(/\[([^\]]+)\]\s*$/);
    var fig = mb ? mb[1].trim() : body.replace(/\s+/g, " ").slice(0, 60);
    if (!fig) return;
    (groups[attr] = groups[attr] || []).push(fig);
  });

  var lines = Object.keys(groups).map(function (attr) {
    // de-dup within a group
    var uniq = groups[attr].filter(function (v, i) { return groups[attr].indexOf(v) === i; });
    return "  " + attr + ": " + uniq.join(" | ");
  });
  if (!lines.length) return "";

  return "\n" +
    "VERIFIED FIGURES — the ONLY numbers, penalties, settlements, and dates allowed on screen:\n" +
    lines.join("\n") + "\n" +
    "RULES: (1) Any figure/penalty/settlement/agency/date NOT listed above is FORBIDDEN on screen — " +
    "the brief and checkpoints may contain unverified numbers; ignore them. " +
    "(2) Keep hedges (\"potentially\", \"approximately\") exactly as written. " +
    "(3) Figures marked \"Company self-reported\" are the bank's OWN disclosure — never present them as a regulator finding. " +
    "(4) Never compute a % change between two figures from different review scopes.\n";
}

function buildScenesPrompt(idea, master, voiceover, narrative, dataMoments, sceneBlueprint) {
  // Accept both old (dataPoints) and new (dataMoments, sceneBlueprint) call signatures
  // Stage4 upgraded.gs calls this with (idea, master, voiceover, narrative, dataMoments, sceneBlueprint)
  dataMoments    = dataMoments    || "";
  sceneBlueprint = sceneBlueprint || "";
  const verifiedFiguresBlock = buildVerifiedFiguresBlock_(idea.id);

  // A scene's ON-SCREEN length IS its Stage 7B narration clip length, so scene
  // duration is decided HERE, by how finely the voiceover is chunked. The first
  // Nissan cut averaged 31s/scene (one scene held 103s = 30% of the film) because
  // this table asked for 15-25s. At ~140wpm a 12s beat is ~28 words, so a ~790-word
  // Standard script is ~28 beats. maxCheckpoints stops one card type dominating
  // (that cut used 7 of 11 scenes as checkpoints).
  const FORMAT_SPEC = {
    "Short (< 90s)"        : { totalDuration: "0:00-1:30",   avgSceneDuration: "7-11 sec",  avgSecs: 9,    minScenes: 8,  maxScenes: 12, maxCheckpoints: 3 },
    "Standard (4-7 min)"   : { totalDuration: "4:00-7:00",   avgSceneDuration: "11-14 sec", avgSecs: 12.5, minScenes: 26, maxScenes: 30, maxCheckpoints: 4 },
    "Deep Dive (10-15 min)": { totalDuration: "10:00-15:00", avgSceneDuration: "12-16 sec", avgSecs: 14,   minScenes: 45, maxScenes: 60, maxCheckpoints: 6 }
  };
  const spec = FORMAT_SPEC[idea.targetFormat] || FORMAT_SPEC["Standard (4-7 min)"];

  // ── Dynamic scene count ───────────────────────────────────────────────────
  const checkpoints = master.checkpoints
    ? master.checkpoints.toString().split("->").concat(
        master.checkpoints.toString().split("\u2192")
      ).map(c => c.trim()).filter(c => c.length > 0)
    : [];
  // Deduplicate (split by both -> and → may double up)
  const uniqueCPs = [...new Set(checkpoints)];
  const checkpointCount = uniqueCPs.length;

  const dataMomentCount = dataMoments ? (dataMoments.match(/DM_\d+:/g) || []).length : 0;
  // Checkpoints are capped so they cannot crowd out every other component.
  const cpScenes    = Math.min(checkpointCount, spec.maxCheckpoints);
  const totalScenes = spec.minScenes;
  // Whatever the checkpoints and the bookends don't cover has to be carried by
  // data/statement/evidence scenes — this is the number that forces variety.
  const otherScenes = Math.max(0, totalScenes - cpScenes - 3);   // 3 = hook + mid-verdict + timeline

  // Runtime now follows from scene count x average dwell (dataScenes is gone —
  // scene count is driven by the format target, not by data-moment count).
  const estimatedSecs   = Math.round(totalScenes * spec.avgSecs);
  const durationDisplay = Math.floor(estimatedSecs / 60) + ":" +
                          String(estimatedSecs % 60).padStart(2, "0") +
                          " (estimated — sync to voiceover length)";

  return `
You are building the Visual Library (scene list) for a GovernX YouTube video.

CONTENT BRIEF:
Company/Topic  : ${idea.company}
Title          : ${master.title}
Discipline     : ${master.discipline}
Checkpoints    : ${master.checkpoints}

TARGET FORMAT     : ${idea.targetFormat}
TOTAL DURATION    : ${durationDisplay}
TOTAL SCENES      : ${spec.minScenes}–${spec.maxScenes} (MANDATORY RANGE — see PACING below)
AVG SCENE DURATION: ${spec.avgSceneDuration}

SCENE BREAKDOWN REQUIRED:
- 1 opening Text card (hook stat — most dramatic number or date)
- AT MOST ${cpScenes} Checkpoint cards (reverse order, outcome first)
- ~${otherScenes} Infographic / Data Table / Text scenes carrying everything else
- 1 mid-film Text verdict card
- 1 final Timeline scene (ALWAYS last)
Target: ${spec.minScenes}–${spec.maxScenes} scenes total

═══════════════════════════════════════════════════════
PACING — THE MOST COMMON FAILURE. READ TWICE.
═══════════════════════════════════════════════════════
A scene stays on screen for exactly as long as the narration you assign to it.
There is no other timing control. So scene length is decided HERE, by how finely
you cut the voiceover.

HARD RULES:
  1. Cut the voiceover into ${spec.minScenes}–${spec.maxScenes} scenes.
     One scene = ONE complete thought = 20–35 spoken words.
  2. NO SCENE MAY EXCEED 58 WORDS (~25 seconds). If a passage is longer,
     SPLIT IT into consecutive scenes. This is not optional.
  3. A passage that ENUMERATES A LIST gets ONE SCENE PER ITEM.
     If the narration says "60 million here, 1.85 billion there, 14.7 million
     there…", that is 3+ scenes, not one.
  4. Match dwell to how much there is to read:
       5–8s  (12–19 words) — single statement, one big number
       8–12s (19–28 words) — checkpoint, one comparison
       12–18s (28–42 words) — KPI dashboard, risk matrix, bar chart, timeline
  5. Never assign the same narration sentence to two scenes.

A REAL FAILURE TO AVOID: a previous film gave ONE scene 103 seconds of narration
listing six separate figures. It rendered as a single poster held for 1 minute
43 seconds — 30% of the film on one static image. That is six scenes.

═══════════════════════════════════════════════════════
VARIETY — NO TYPE MAY DOMINATE
═══════════════════════════════════════════════════════
  • Use AT LEAST 8 DIFFERENT scene/infographic variants across the film.
  • NO variant may appear more than 4 times.
  • Checkpoint cards: MAXIMUM ${cpScenes}. (A previous film used 7 of 11 scenes
    as checkpoints and looked like the same card repeating.)
  • NEVER put two scenes of the same variant back to back.
  • Open on a hook statement. Close on the Timeline. Never open or close on a
    checkpoint.

═══════════════════════════════════════════════════════
PRODUCTION MODEL: REMOTION ONLY — DATA-DRIVEN PRESENTATION
═══════════════════════════════════════════════════════
GovernX is a data-driven presentation channel.
NO stock footage. NO AI-generated video. NO B-roll. NO CapCut builds.

EVERY scene is built in Remotion. Source = Remotion for ALL scenes.

Valid scene types and their Remotion compositions:
  Text        → Remotion: TextImpactScene
  Infographic → Remotion: InfographicScene
  Data Table  → Remotion: InfographicScene (BEFORE_AFTER_CARD variant)
  Checkpoint  → Remotion: CheckpointCard
  Timeline    → Remotion: TimelineReveal

BUILT_WHERE must always be: Remotion — [CompositionName]
SCENE_SOURCE must always be: Remotion
LICENSE must always be: Original

INFOGRAPHIC VARIANTS — use the most powerful for each data moment.
All nine render today. Spread the work across them; do not lean on two.
  [INFOGRAPHIC] LINE_GRAPH        → metric over time (3+ points, same variable)
  [INFOGRAPHIC] SPLIT_COMPARISON  → two entities side-by-side (before/after, A vs B)
  [INFOGRAPHIC] DATA_CALLOUT      → single powerful statistic
  [INFOGRAPHIC] COUNTER_ANIMATION → counting up to one dramatic figure
  [INFOGRAPHIC] BEFORE_AFTER_CARD → governance structure: as-designed vs as-run
  [INFOGRAPHIC] BAR_CHART         → comparing magnitudes across 3–6 items
                                    (penalties, settlements, outlays by category)
  [INFOGRAPHIC] KPI_DASHBOARD     → 3–4 headline figures at once, each attributed
                                    (the "case in numbers" beat)
  [INFOGRAPHIC] RISK_MATRIX       → where the control gap lived
                                    (likelihood x impact, 1=low 2=med 3=high)
  [INFOGRAPHIC] PROGRESS_GAUGE    → coverage / completeness as a percentage
                                    (independent directors, audit coverage)

The last four are UNDER-USED. If the evidence supports a KPI_DASHBOARD,
RISK_MATRIX, BAR_CHART or PROGRESS_GAUGE, choosing a plain DATA_CALLOUT instead
is a scoring failure. Never invent numbers to fill one — every figure must come
from the verified figures block below.

SCENE SEQUENCING RULES:
  1. Scene 1: the hook — an OPENING_HOOK or a shatter Text card on the strongest number.
  2. Checkpoints in reverse-engineering order (outcome first, root cause last).
  3. Between each checkpoint: at least one Infographic / data-viz / case-file scene.
  4. The Timeline is a RECAP, not the ending. Place it near the close (a chronology
     of the whole arc), but it must NOT be the final scene — a film that ends on a
     timeline ends on a reference table, which is anticlimactic.
  5. CLOSE ON THE THESIS. The final scene is the verdict / thesis statement
     (VERDICT_CARD or a verdict Text card) — the sentence the whole film was
     building toward. If there is a Call to Action, it belongs on this final
     scene, spoken ONCE.
  Typical strong ending, in order:
     … → Timeline (recap) → DATA_WALL or EVIDENCE_CARD (evidence on the record)
       → VERDICT_CARD / thesis statement (the close, CTA here).

⚠ ONE SENTENCE = ONE SCENE. NEVER assign the same VOICEOVER_SYNC to two scenes.
  If you are running out of narration near the end, use FEWER scenes — do not pad
  by repeating a sentence across several cards. A previous film spoke its closing
  two sentences three times each across five scenes; that is a scoring failure.
  Every scene's VOICEOVER_SYNC must be a DISTINCT span of the script, in order,
  with no sentence reused. It is correct to emit fewer than the target scene count
  when the remaining script cannot fill more scenes without repetition.
${verifiedFiguresBlock}
═══════════════════════════════════════════════════════
GOVERNX CHECKPOINT CARDS (${checkpointCount} required)
═══════════════════════════════════════════════════════
${uniqueCPs.map((cp, i) => "  " + (i + 1) + ". " + cp).join("\n")}
Note: the checkpoint list above is from the pre-verification brief. Rewrite any
checkpoint whose figures are not in VERIFIED EVIDENCE to use only cleared facts.

Every Checkpoint scene:
- SCENE_SOURCE: Remotion
- BUILT_WHERE: Remotion — CheckpointCard
- CHECKPOINT_DATE: specific year or date
- CHECKPOINT_EVENT: one sharp sentence — max 12 words
- CHECKPOINT_ANGLE: governance/process failure in CAPS — max 10 words
- ASSEMBLY_NOTES: date=... | event=... | angle=... | variant=standard (use variant=root for ROOT CAUSE, variant=outcome for first checkpoint)

TIMELINE RECAP SCENE (place near the close, NOT as the final scene — see rule 4):
- SCENE_SOURCE: Remotion
- BUILT_WHERE: Remotion — TimelineReveal
- DESCRIPTION: all checkpoints listed with arrows
- ASSEMBLY_NOTES: checkpoints=[comma-separated dates in order] | variant=standard

${sceneBlueprint ? `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SCENE BLUEPRINT — EXECUTE THIS EXACTLY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Stage 3 wrote the full scene plan. Translate each SCENE_BP_ into a SCENE_ output block.
Do NOT add, remove, or change scene types. Map directly:
  Text       → SCENE_TYPE: Text       | SOURCE: Remotion | BUILT_WHERE: Remotion — TextImpactScene
  Infographic→ SCENE_TYPE: Infographic| SOURCE: Remotion | BUILT_WHERE: Remotion — InfographicScene
  Data Table → SCENE_TYPE: Data Table | SOURCE: Remotion | BUILT_WHERE: Remotion — InfographicScene
  Checkpoint → SCENE_TYPE: Checkpoint | SOURCE: Remotion | BUILT_WHERE: Remotion — CheckpointCard
  Timeline   → SCENE_TYPE: Timeline   | SOURCE: Remotion | BUILT_WHERE: Remotion — TimelineReveal

${sceneBlueprint}

Infographic DESCRIPTION format:
  LINE_GRAPH:        [INFOGRAPHIC] LINE_GRAPH — [label] | [points: year:value,year:value]
  SPLIT_COMPARISON:  [INFOGRAPHIC] SPLIT_COMPARISON — [left] vs [right] | [data pairs]
  DATA_CALLOUT:      [INFOGRAPHIC] DATA_CALLOUT — [value] | [label] | [context]
  COUNTER_ANIMATION: [INFOGRAPHIC] COUNTER_ANIMATION — [from]->[to] | [unit] | [label]
  BEFORE_AFTER_CARD: [INFOGRAPHIC] BEFORE_AFTER_CARD — [before rows] | vs | [after rows]
` : dataMoments ? `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DATA MOMENTS — INFOGRAPHIC SCENES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
No scene blueprint available. Use DATA_MOMENTS to generate Infographic scenes.
For each DM_ entry generate one dedicated Infographic or Data Table scene.

${dataMoments}

DATA_TYPE mapping:
  TIME_SERIES  → [INFOGRAPHIC] LINE_GRAPH
  COMPARISON   → [INFOGRAPHIC] SPLIT_COMPARISON
  SINGLE_STAT  → [INFOGRAPHIC] DATA_CALLOUT
  COUNTER      → [INFOGRAPHIC] COUNTER_ANIMATION
  BEFORE_AFTER → [INFOGRAPHIC] BEFORE_AFTER_CARD
` : ""}

NARRATIVE FLOW:
${narrative}

VOICEOVER SCRIPT (first 600 chars):
${voiceover.substring(0, 600)}...

VALID VALUES:
Scene Types : Text | Infographic | Data Table | Checkpoint | Timeline
Sources     : Remotion
Licenses    : Original
Status      : Needed

Return each scene in EXACTLY this format:

SCENE_1_START
SCENE_TYPE: [value — Text | Infographic | Data Table | Checkpoint | Timeline]
DESCRIPTION: [for Infographic: start with [INFOGRAPHIC] VARIANT_NAME — then exact data with years and labels]
SCENE_SOURCE: Remotion
SCENE_LINK: Remotion render — see BUILT_WHERE
SCENE_TIMESTAMP: [MM:SS]
LICENSE: Original
SCENE_STATUS: Needed
BUILT_WHERE: Remotion — [TextImpactScene | InfographicScene | CheckpointCard | TimelineReveal]
ASSEMBLY_NOTES: [structured Remotion data — pipe-separated key=value pairs]
VEO_PROMPT: N/A — built in Remotion
CHECKPOINT_DATE: [for Checkpoint scenes ONLY — write NONE for all other types]
CHECKPOINT_EVENT: [for Checkpoint scenes ONLY — write NONE for all other types]
CHECKPOINT_ANGLE: [for Checkpoint scenes ONLY — write NONE for all other types]
SCENE_1_END

[repeat for all scenes — final scene MUST be SCENE_TYPE: Timeline]

ASSEMBLY_NOTES FORMAT BY SCENE TYPE:
- Text:        mainText="[primary]" | subText="[secondary]" | type=[shatter|verdict|default]
- Infographic: type=[VARIANT] | [variant-specific key=value pairs — exact numbers, years, labels]
- Checkpoint:  date=[year] | event=[sentence] | angle=[CAPS] | variant=[standard|root|outcome]
- Timeline:    checkpoints=[date1,date2,...] | variant=standard
- Data Table:  type=BEFORE_AFTER_CARD | before_label=... | before_rows=... | after_label=... | after_rows=...

${VISUAL_INTELLIGENCE_SKILL}
`;}

function writeScenes(contentId, raw) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.VISUAL);

  // ── Upsert: delete any existing scene rows for this contentId ─────────────
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][COL_VISUAL.ID - 1].toString().trim() === contentId) {
      rowsToDelete.push(i + 1);
    }
  }
  if (rowsToDelete.length > 0) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Scenes Already Exist",
      "Found " + rowsToDelete.length + " existing scenes for: " + contentId + "\n\nOverwrite with new generation?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    rowsToDelete.forEach(row => sheet.deleteRow(row));
  }

  const sceneBlocks = raw.match(/SCENE_\d+_START([\s\S]*?)SCENE_\d+_END/g) || [];

  sceneBlocks.forEach((block, index) => {

    // ── Robust getValue: captures value on same line only, returns "" if blank ──
    const getValue = (field) => {
      const match = block.match(new RegExp(field + ":\\s*([^\\n\\r]*)"));
      if (!match) return "";
      const val = match[1].trim();
      // Reject if value looks like a field name (ALL_CAPS with colon pattern)
      if (/^[A-Z_]{3,}:/.test(val)) return "";
      // Reject known sentinel/empty values
      if (val === "SCENE_1_END" || val.startsWith("SCENE_") ||
          val === "NONE" || val === "" ) return "";
      return val;
    };

    const lastRow = sheet.getLastRow() + 1;

    // Clear validation on free-text and non-standard value fields before writing
    [COL_VISUAL.DESCRIPTION, COL_VISUAL.LINK,
     COL_VISUAL.SOURCE, COL_VISUAL.SCENE_TYPE,
     COL_VISUAL.BUILT_WHERE, COL_VISUAL.ASSEMBLY_NOTES,
     COL_VISUAL.VEO_PROMPT, COL_VISUAL.CHECKPOINT_DATE,
     COL_VISUAL.CHECKPOINT_EVENT, COL_VISUAL.CHECKPOINT_ANGLE].forEach(col => {
      sheet.getRange(lastRow, col).clearDataValidations();
    });

    sheet.getRange(lastRow, COL_VISUAL.ID              ).setValue(contentId);
    sheet.getRange(lastRow, COL_VISUAL.SCENE_NUM       ).setValue(index + 1);
    sheet.getRange(lastRow, COL_VISUAL.SCENE_TYPE      ).setValue(getValue("SCENE_TYPE"));
    sheet.getRange(lastRow, COL_VISUAL.DESCRIPTION     ).setValue(getValue("DESCRIPTION"));
    sheet.getRange(lastRow, COL_VISUAL.SOURCE          ).setValue(getValue("SCENE_SOURCE"));
    sheet.getRange(lastRow, COL_VISUAL.LINK            ).setValue(getValue("SCENE_LINK"));
    sheet.getRange(lastRow, COL_VISUAL.TIMESTAMP       ).setValue(getValue("SCENE_TIMESTAMP"));
    sheet.getRange(lastRow, COL_VISUAL.LICENSE         ).setValue(getValue("LICENSE"));
    sheet.getRange(lastRow, COL_VISUAL.STATUS          ).setValue(getValue("SCENE_STATUS"));
    sheet.getRange(lastRow, COL_VISUAL.BUILT_WHERE     ).setValue(getValue("BUILT_WHERE"));
    sheet.getRange(lastRow, COL_VISUAL.ASSEMBLY_NOTES  ).setValue(getValue("ASSEMBLY_NOTES"));
    sheet.getRange(lastRow, COL_VISUAL.VEO_PROMPT      ).setValue(getValue("VEO_PROMPT"));
    sheet.getRange(lastRow, COL_VISUAL.CHECKPOINT_DATE ).setValue(getValue("CHECKPOINT_DATE"));
    sheet.getRange(lastRow, COL_VISUAL.CHECKPOINT_EVENT).setValue(getValue("CHECKPOINT_EVENT"));
    sheet.getRange(lastRow, COL_VISUAL.CHECKPOINT_ANGLE).setValue(getValue("CHECKPOINT_ANGLE"));
  });
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 5 — Create Publishing Tracker Row
// ════════════════════════════════════════════════════════════════════════════════
function createPublishingRow() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const master = getMasterContent(idea.id);
  const ui     = SpreadsheetApp.getUi();

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET.PUBLISHING);

    // Check if row already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
        ui.alert(
          "Already Exists",
          "A Publishing Tracker row already exists for: " + idea.id,
          ui.ButtonSet.OK
        );
        return;
      }
    }

    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, COL_PUBLISHING.ID         ).setValue(idea.id);
    sheet.getRange(lastRow, COL_PUBLISHING.TITLE_FINAL).setValue(master ? master.title : idea.company);

    // Add benchmark notes
    sheet.getRange(lastRow, COL_PUBLISHING.NOTES).setValue(
      "Benchmarks: CTR Scale >" + BENCHMARKS.CTR_SCALE + "% | " +
      "Retention Scale >" + BENCHMARKS.RETENTION_SCALE + "% | " +
      "Views Kill <" + BENCHMARKS.VIEWS_KILL + " at 30 days"
    );

    updatePipelineStatus_(idea.id, "S5", "✅");
    ui.alert(
      "✅ Stage 5 Complete",
      "Publishing Tracker row created for: " + idea.id + "\n\n" +
      "Next: Stage 6 — Export Production Package,\n" +
      "then Stage 7 / 7B — Voiceover Audio.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S5", "❌");
    logError("Stage 5 — Publishing", idea.id, "Runtime Error", err.message);
    ui.alert("❌ Stage 5 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 6 — Export Production Package
// Generates a self-contained text file with everything needed to produce the video
// Format: structured prompt ready to paste into CapCut / KlingAI / any video AI tool
// ════════════════════════════════════════════════════════════════════════════════
function exportProductionPackage() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);

  // ── Pull script data ──────────────────────────────────────────────────────
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let script = null;
  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      script = {
        hook        : scriptData[i][COL_SCRIPT.HOOK              - 1],
        narrative   : scriptData[i][COL_SCRIPT.NARRATIVE_FLOW    - 1],
        voiceover   : scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT  - 1],
        closing     : scriptData[i][COL_SCRIPT.GRC_BPR_CLOSING   - 1],
        format      : scriptData[i][COL_SCRIPT.TARGET_FORMAT     - 1],
        cta         : scriptData[i][COL_SCRIPT.CALL_TO_ACTION    - 1],
        sections    : scriptData[i][COL_SCRIPT.SECTIONS          - 1]
      };
      break;
    }
  }

  if (!script) {
    ui.alert("No script found for: " + idea.id + "\nRun Stage 3 first.");
    return;
  }

  // ── Pull scene data ───────────────────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  let scenes = [];
  if (visualSheet) {
    const visualData = visualSheet.getDataRange().getValues();
    for (let i = 1; i < visualData.length; i++) {
      if (visualData[i][COL_VISUAL.ID - 1].toString().trim() === idea.id) {
        scenes.push({
          num         : visualData[i][COL_VISUAL.SCENE_NUM   - 1],
          type        : visualData[i][COL_VISUAL.SCENE_TYPE  - 1],
          description : visualData[i][COL_VISUAL.DESCRIPTION - 1],
          source      : visualData[i][COL_VISUAL.SOURCE      - 1],
          link        : visualData[i][COL_VISUAL.LINK        - 1],
          timestamp   : visualData[i][COL_VISUAL.TIMESTAMP   - 1],
          license     : visualData[i][COL_VISUAL.LICENSE     - 1]
        });
      }
    }
  }

  // ── Pull research sources used ────────────────────────────────────────────
  const researchSheet = ss.getSheetByName(SHEET.RESEARCH);
  let usedSources = [];
  if (researchSheet) {
    const resData = researchSheet.getDataRange().getValues();
    for (let i = 1; i < resData.length; i++) {
      if (resData[i][COL_RESEARCH.ID - 1].toString().trim() === idea.id &&
          resData[i][COL_RESEARCH.USED_IN_SCRIPT - 1].toString().trim() === "YES") {
        usedSources.push({
          type    : resData[i][COL_RESEARCH.SOURCE_TYPE - 1],
          details : resData[i][COL_RESEARCH.DETAILS     - 1],
          insight : resData[i][COL_RESEARCH.KEY_INSIGHT - 1]
        });
      }
    }
  }

  // ── Build the export document ─────────────────────────────────────────────
  const divider = "═".repeat(60);
  const line    = "─".repeat(60);

  let doc = "";

  doc += divider + "\n";
  doc += "GOVERNX — VIDEO PRODUCTION PACKAGE\n";
  doc += divider + "\n";
  doc += "Content ID    : " + idea.id + "\n";
  doc += "Title         : " + (master ? master.title : idea.company) + "\n";
  doc += "Company/Topic : " + idea.company + "\n";
  doc += "Discipline    : " + (master ? master.discipline : "") + "\n";
  doc += "Format        : " + script.format + "\n";
  doc += "Language      : " + idea.language + "\n";
  doc += "Series        : " + idea.series + "\n";
  doc += "Exported      : " + new Date().toLocaleString() + "\n";
  doc += divider + "\n\n";

  // ── SECTION 1: VOICEOVER SCRIPT ───────────────────────────────────────────
  doc += "SECTION 1 — VOICEOVER SCRIPT\n";
  doc += line + "\n";
  doc += "HOOK:\n" + script.hook + "\n\n";
  doc += "FULL SCRIPT:\n" + script.voiceover + "\n\n";
  doc += "GRC/BPR CLOSING ARGUMENT:\n" + script.closing + "\n\n";
  doc += "CALL TO ACTION: " + script.cta + "\n\n";

  // ── SECTION 2: SCENE LIST ─────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 2 — SCENE LIST (" + scenes.length + " scenes)\n";
  doc += line + "\n";
  doc += "FORMAT FOR VIDEO AI TOOL:\n";
  doc += "Paste each scene description as a prompt. Timestamp = placement in video.\n\n";

  if (scenes.length > 0) {
    scenes.forEach(scene => {
      doc += "SCENE " + scene.num + " [" + scene.timestamp + "] — " + scene.type.toUpperCase() + "\n";
      doc += "Visual: " + scene.description + "\n";
      doc += "Source: " + scene.source + (scene.link ? " → " + scene.link : "") + "\n";
      doc += "License: " + scene.license + "\n";
      doc += "\n";
    });
  } else {
    doc += "(No scenes generated yet — run Stage 4 first)\n\n";
  }

  // ── SECTION 3: AI VIDEO TOOL PROMPT ──────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 3 — AI VIDEO TOOL PROMPT\n";
  doc += line + "\n";
  doc += "Copy and paste this directly into CapCut AI, KlingAI, or similar tools:\n\n";
  doc += "---\n";
  doc += "Create a " + script.format + " video titled: \"" + (master ? master.title : idea.company) + "\"\n\n";
  doc += "NARRATIVE STRUCTURE:\n" + script.narrative + "\n\n";
  doc += "VOICEOVER (read exactly as written):\n" + script.voiceover + "\n\n";
  doc += "VISUAL STYLE: Documentary style. Sharp cuts. Text overlays for key facts. ";
  doc += "Dark professional tone. No stock-photo feel. Data visualizations where applicable.\n\n";
  if (scenes.length > 0) {
    doc += "SCENE BREAKDOWN:\n";
    scenes.forEach(scene => {
      doc += "[" + scene.timestamp + "] " + scene.description + "\n";
    });
  }
  doc += "---\n\n";

  // ── SECTION 4: EVIDENCE TRAIL ────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 4 — EVIDENCE TRAIL (" + usedSources.length + " sources used in script)\n";
  doc += line + "\n";
  if (usedSources.length > 0) {
    usedSources.forEach((src, i) => {
      doc += (i + 1) + ". [" + src.type + "] " + src.details + "\n";
      doc += "   Key Insight: " + src.insight + "\n\n";
    });
  } else {
    doc += "(Mark sources as Used in Script? = YES in Research Database to include them here)\n\n";
  }

  doc += divider + "\n";
  doc += "END OF PRODUCTION PACKAGE — " + idea.id + "\n";
  doc += divider + "\n";

  // ── Write to a Google Doc inside content subfolder ───────────────────────
  try {
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
    const docTitle      = idea.id + " — " + idea.company + " — Production Package";

    // Delete existing doc if any
    const existingFiles = contentFolder.getFilesByName(docTitle);
    while (existingFiles.hasNext()) existingFiles.next().setTrashed(true);

    // Create doc, move to content folder
    const newDoc = DocumentApp.create(docTitle);
    newDoc.getBody().setText(doc);
    newDoc.saveAndClose();

    const file = DriveApp.getFileById(newDoc.getId());
    contentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const docUrl = newDoc.getUrl();

    // ── Write URL to Publishing Tracker ──────────────────────────────────
    writePublishingLink(idea.id, idea.company, master,
      COL_PUBLISHING.PRODUCTION_PACKAGE, docUrl, "📄 Open Package");

    updatePipelineStatus_(idea.id, "S6", "✅");
    ui.alert(
      "✅ Stage 6 Complete — Production Package Exported",
      "Saved to Google Drive:\n" +
      "📁 " + DRIVE_FOLDER_NAME + " / " + idea.id + "\n" +
      "📄 " + docTitle + "\n\n" +
      "Contains:\n" +
      "• Full voiceover script\n" +
      "• Scene list (" + scenes.length + " scenes)\n" +
      "• Ready-to-paste AI video tool prompt\n" +
      "• Evidence trail (" + usedSources.length + " sources)\n\n" +
      "Link saved in Publishing Tracker → Production Package column.\n\n" +
      docUrl,
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S6", "❌");
    logError("Stage 6 — Export", idea.id, "Google Doc creation failed", err.message);
    ui.alert("❌ Stage 6 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// SHARED HELPERS — Folder management & Publishing Tracker link writer
// ════════════════════════════════════════════════════════════════════════════════

// Returns (or creates) the content subfolder: GovernX Production Packages / [ID]
function getOrCreateContentFolder(contentId, company) {
  // Get or create root folder
  let rootFolder;
  const rootSearch = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  rootFolder = rootSearch.hasNext()
    ? rootSearch.next()
    : DriveApp.createFolder(DRIVE_FOLDER_NAME);

  // Get or create content subfolder named by ID
  const subName   = contentId + " — " + company;
  const subSearch = rootFolder.getFoldersByName(subName);
  return subSearch.hasNext()
    ? subSearch.next()
    : rootFolder.createFolder(subName);
}

// Returns (or creates) the Scenes subfolder inside the content folder
function getOrCreateScenesFolder(contentId, company) {
  const contentFolder = getOrCreateContentFolder(contentId, company);
  const sceneSearch   = contentFolder.getFoldersByName(DRIVE_SCENES_SUBFOLDER);
  return sceneSearch.hasNext()
    ? sceneSearch.next()
    : contentFolder.createFolder(DRIVE_SCENES_SUBFOLDER);
}


// ════════════════════════════════════════════════════════════════════════════════
// STAGE 6 TOOL — Cleanup Scenes Folder
// Trashes all intermediate files (scene clips + per-scene audio) from the
// Drive Scenes subfolder after the final video has been uploaded to YouTube.
// The final assembled video (contentFolder/*_final_video.mp4) is NOT touched.
// ════════════════════════════════════════════════════════════════════════════════
function cleanupScenesFolder() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();

  // ── Locate Scenes subfolder ───────────────────────────────────────────────
  const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
  const sceneSearch   = contentFolder.getFoldersByName(DRIVE_SCENES_SUBFOLDER);
  if (!sceneSearch.hasNext()) {
    ui.alert("Scenes Folder Not Found",
      "No Scenes subfolder exists yet for: " + idea.id,
      ui.ButtonSet.OK);
    return;
  }
  const scenesFolder = sceneSearch.next();

  // ── Collect file names for preview ───────────────────────────────────────
  const fileIter = scenesFolder.getFiles();
  const files    = [];
  while (fileIter.hasNext()) {
    const f = fileIter.next();
    files.push({ id: f.getId(), name: f.getName(), size: f.getSize() });
  }

  if (files.length === 0) {
    ui.alert("✅ Scenes Folder Already Empty",
      "No files to clean up in: " + DRIVE_SCENES_SUBFOLDER,
      ui.ButtonSet.OK);
    return;
  }

  const totalMB = (files.reduce(function(s, f) { return s + f.size; }, 0) / 1024 / 1024).toFixed(1);

  // Show what will be trashed (up to 10 names to keep dialog readable)
  const preview = files.slice(0, 10).map(function(f) {
    return "  • " + f.name;
  }).join("\n") + (files.length > 10 ? "\n  … and " + (files.length - 10) + " more" : "");

  const confirm = ui.alert(
    "🗂️ Cleanup Scenes Folder",
    "Found " + files.length + " file(s) (" + totalMB + " MB) in the Scenes folder for: " + idea.id + "\n\n" +
    preview + "\n\n" +
    "⚠️  These will be MOVED TO TRASH (Drive trash — recoverable for 30 days).\n\n" +
    "The final assembled video (*_final_video.mp4) in the content root is NOT affected.\n\n" +
    "Proceed with cleanup?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  // ── Trash all files in Scenes subfolder ───────────────────────────────────
  let trashed = 0;
  let failed  = 0;
  files.forEach(function(f) {
    try {
      DriveApp.getFileById(f.id).setTrashed(true);
      trashed++;
    } catch (e) {
      failed++;
      Logger.log("Cleanup: failed to trash " + f.name + ": " + e.message);
    }
  });

  Logger.log("Cleanup: " + trashed + " files trashed from " + idea.id + " Scenes folder.");

  ui.alert(
    "✅ Cleanup Complete",
    "Moved " + trashed + " file(s) to Drive trash.\n" +
    (failed > 0 ? "⚠️ " + failed + " file(s) could not be trashed (see Logs).\n" : "") +
    "\nSpace freed: ~" + totalMB + " MB\n\n" +
    "Files are recoverable from Drive trash for 30 days.",
    ui.ButtonSet.OK
  );
}

// Writes a hyperlink to a specific Publishing Tracker column for a content ID
function writePublishingLink(contentId, company, master, column, url, label) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) return;

  const pubData = pubSheet.getDataRange().getValues();
  let   found   = false;

  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === contentId) {
      pubSheet.getRange(i + 1, column).clearDataValidations();
      pubSheet.getRange(i + 1, column)
        .setFormula('=HYPERLINK("' + url + '","' + label + '")');
      found = true;
      break;
    }
  }

  // Auto-create Publishing row if missing
  if (!found) {
    const lastRow = pubSheet.getLastRow() + 1;
    // Clear validation on all columns we'll write to
    pubSheet.getRange(lastRow, COL_PUBLISHING.ID         ).clearDataValidations();
    pubSheet.getRange(lastRow, COL_PUBLISHING.TITLE_FINAL).clearDataValidations();
    pubSheet.getRange(lastRow, column                    ).clearDataValidations();

    pubSheet.getRange(lastRow, COL_PUBLISHING.ID         ).setValue(contentId);
    pubSheet.getRange(lastRow, COL_PUBLISHING.TITLE_FINAL).setValue(
      master ? master.title : company
    );
    pubSheet.getRange(lastRow, column)
      .setFormula('=HYPERLINK("' + url + '","' + label + '")');
  }
}

// ── ElevenLabs voice ID selector ─────────────────────────────────────────────
// Returns the correct voice ID based on content language flag.
// Arabic / Bilingual → ELEVENLABS_VOICE_ID_AR  (Arabic-native voice)
// English            → ELEVENLABS_VOICE_ID_EN  (English authoritative voice)
// Falls back to legacy ELEVENLABS_VOICE_ID if per-language keys are not set.
function getElevenLabsVoiceId_(props, language) {
  const isArabic = (language === "Arabic" || language === "Bilingual");
  const key      = isArabic ? "ELEVENLABS_VOICE_ID_AR" : "ELEVENLABS_VOICE_ID_EN";
  const voiceId  = props.getProperty(key);
  if (voiceId) return { id: voiceId, key: key };

  // Legacy fallback — single voice ID
  const legacy = props.getProperty("ELEVENLABS_VOICE_ID");
  if (legacy) return { id: legacy, key: "ELEVENLABS_VOICE_ID (legacy fallback)" };

  return null;
}


/* ── Make figures speak naturally in ElevenLabs ────────────────────────────────
   VOICEOVER_SYNC is written for the screen ("$140 million", "¥320 million",
   "36%"). ElevenLabs reads a LEADING "$" awkwardly ("140 dollars million") and
   does not know "¥"/"€" at all — so it drops or mangles them ("140 m"). Move the
   currency to a spoken word AFTER the amount and expand % and magnitude letters.
   The on-screen REMOTION_DATA is NOT touched — this only shapes the audio text.
   Dates ("November 19, 2018"), plain "billion yen" figures, and years are left
   alone. Applied at both Stage 7 (full VO) and Stage 7B (per-scene). */
function speakNumbers_(text) {
  const CUR = { "$": "dollars", "¥": "yen", "€": "euros", "£": "pounds" };
  const MAG = { k: "thousand", m: "million", b: "billion", t: "trillion" };
  const ORD = { "1": "first", "2": "second", "3": "third", "4": "fourth" };
  let t = String(text || "");

  // quarters: Q4 → "the fourth quarter"; Q4 2023 → "the fourth quarter of 2023"
  t = t.replace(/\bQ([1-4])(?:\s+(\d{4}))?\b/g,
    function (m, q, yr) { return "the " + ORD[q] + " quarter" + (yr ? " of " + yr : ""); });

  // currency symbol + amount (+ optional magnitude word or letter) → amount magnitude currency
  t = t.replace(/([$¥€£])\s?(\d[\d.,]*)(?:\s?(million|billion|trillion|thousand|[KMBT]))?\b/gi,
    function (m, sym, num, mag) {
      const magWord = mag ? " " + (MAG[mag.toLowerCase()] || mag.toLowerCase()) : "";
      return num + magWord + " " + CUR[sym];
    });

  // compact magnitude letter attached to a bare number (no currency): 9.078B, 140M
  t = t.replace(/\b(\d[\d.,]*)([KMBT])\b/g,
    function (m, num, suf) { return num + " " + MAG[suf.toLowerCase()]; });

  // percent
  t = t.replace(/(\d[\d.,]*)\s?%/g, function (m, num) { return num + " percent"; });

  // multipliers: 2x, 3.5x, 2× → "2 times"
  t = t.replace(/\b(\d[\d.,]*)\s?[x×]\b/gi, function (m, num) { return num + " times"; });

  // numeric range with a dash → "to" (years, plain numbers). Fires only when BOTH
  // sides are numbers, so "10-year" (compound adjective) is deliberately left alone.
  t = t.replace(/(\d[\d.,]*)\s*[–—-]\s*(\d[\d.,]*)/g, function (m, a, b) { return a + " to " + b; });

  return t;
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 7 — Generate Voiceover Audio (ElevenLabs)
// Reads Voiceover Script from Script Bank → generates MP3 → saves to Drive
// ════════════════════════════════════════════════════════════════════════════════
function generateVoiceover() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);

  // Get voiceover script
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let   voiceover   = "";

  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      voiceover = scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1].toString();
      break;
    }
  }

  if (!voiceover || voiceover.trim() === "") {
    ui.alert("No voiceover script found for: " + idea.id + "\nRun Stage 3 first.");
    return;
  }

  // ── Get credentials + language-based voice ID ─────────────────────────────
  const props       = PropertiesService.getScriptProperties();
  const apiKey      = props.getProperty("ELEVENLABS_API_KEY");
  const voiceResult = getElevenLabsVoiceId_(props, idea.language);

  if (!apiKey) {
    ui.alert("ELEVENLABS_API_KEY missing from Script Properties."); return;
  }
  if (!voiceResult) {
    ui.alert(
      "⚠️ ElevenLabs Voice ID Missing",
      "Add the appropriate key to Script Properties:\n\n" +
      "ELEVENLABS_VOICE_ID_EN  — for English content\n" +
      "ELEVENLABS_VOICE_ID_AR  — for Arabic / Bilingual content\n\n" +
      "This content is: " + (idea.language || "English"),
      ui.ButtonSet.OK
    );
    return;
  }

  const voiceId = voiceResult.id;
  Logger.log("Stage 7: language=" + idea.language + " → voice key=" + voiceResult.key + " id=" + voiceId);

  ui.alert(
    "Stage 7 — Generating Voiceover",
    "ElevenLabs is generating the audio for: " + idea.company +
    "\n\nLanguage : " + (idea.language || "English") +
    "\nVoice key: " + voiceResult.key +
    "\n\nThis may take 15–30 seconds depending on script length.",
    ui.ButtonSet.OK
  );

  try {
    // ── Call ElevenLabs TTS API ───────────────────────────────────────────
    const payload = {
      text          : speakNumbers_(voiceover),
      model_id      : ELEVENLABS_MODEL,
      voice_settings: {
        stability        : 0.5,
        similarity_boost : 0.75,
        style            : 0.3,
        use_speaker_boost: true
      }
    };

    const response = UrlFetchApp.fetch(
      ELEVENLABS_API_URL + "/" + voiceId + "?output_format=mp3_44100_128",
      {
        method            : "post",
        contentType       : "application/json",
        headers           : { "xi-api-key": apiKey },
        payload           : JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const code = response.getResponseCode();
    if (code !== 200) {
      throw new Error("ElevenLabs API error " + code + ": " + response.getContentText());
    }

    // ── Save MP3 to Drive subfolder ───────────────────────────────────────
    const audioBlob  = response.getBlob().setContentType("audio/mpeg");
    const fileName   = idea.id + " — " + idea.company + " — Voiceover.mp3";
    const folder     = getOrCreateContentFolder(idea.id, idea.company);

    // Delete existing audio file if any
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const audioFile = folder.createFile(audioBlob.setName(fileName));
    const audioUrl  = "https://drive.google.com/file/d/" + audioFile.getId() + "/view";

    // ── Write link to Publishing Tracker ─────────────────────────────────
    writePublishingLink(idea.id, idea.company, master,
      COL_PUBLISHING.VOICEOVER_AUDIO, audioUrl, "🎵 Play Audio");

    updatePipelineStatus_(idea.id, "S7", "✅");
    ui.alert(
      "✅ Stage 7 Complete — Voiceover Generated",
      "Audio saved to Google Drive:\n" +
      "📁 " + DRIVE_FOLDER_NAME + " / " + idea.id + " — " + idea.company + "\n" +
      "🎵 " + fileName + "\n\n" +
      "Language : " + (idea.language || "English") + "\n" +
      "Voice    : " + voiceResult.key + "\n\n" +
      "Link saved in Publishing Tracker → Voiceover Audio column.\n\n" +
      audioUrl,
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S7", "❌");
    logError("Stage 7 — Voiceover", idea.id, "ElevenLabs API Error", err.message);
    ui.alert("❌ Stage 7 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 7B — Generate Per-Scene Voiceover Audio (ElevenLabs)
// Reads VOICEOVER_SYNC from each Visual Library scene → individual MP3 per scene
// Saves to Drive Scenes subfolder → writes URL back to col 23 (VOICEOVER_AUDIO_URL)
// Used by Shotstack to sync per-scene audio precisely to rendered MP4 clips
// ════════════════════════════════════════════════════════════════════════════════
function generateSceneVoiceovers() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const ui          = SpreadsheetApp.getUi();
  const props       = PropertiesService.getScriptProperties();
  const apiKey      = props.getProperty("ELEVENLABS_API_KEY");
  const voiceResult = getElevenLabsVoiceId_(props, idea.language);

  if (!apiKey) {
    ui.alert("ELEVENLABS_API_KEY missing from Script Properties."); return;
  }
  if (!voiceResult) {
    ui.alert(
      "⚠️ ElevenLabs Voice ID Missing",
      "Add the appropriate key to Script Properties:\n\n" +
      "ELEVENLABS_VOICE_ID_EN  — for English content\n" +
      "ELEVENLABS_VOICE_ID_AR  — for Arabic / Bilingual content\n\n" +
      "This content is: " + (idea.language || "English"),
      ui.ButtonSet.OK
    );
    return;
  }

  const voiceId = voiceResult.id;
  Logger.log("Stage 7B: language=" + idea.language + " → voice key=" + voiceResult.key + " id=" + voiceId);

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();

  // ── Collect scenes that have VOICEOVER_SYNC text ───────────────────────────
  const scenes = [];
  for (let i = 1; i < visualData.length; i++) {
    const row = visualData[i];
    if (row[COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const voiceoverSync = row.length > 20 ? row[COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1].toString().trim() : "";
    if (!voiceoverSync) continue;
    scenes.push({
      row         : i + 1,
      sceneNum    : row[COL_VISUAL.SCENE_NUM  - 1].toString().trim(),
      sceneType   : row[COL_VISUAL.SCENE_TYPE - 1].toString().trim(),
      voiceoverSync,
      existingUrl : row.length > 22 ? row[COL_VISUAL_EXTENDED.VOICEOVER_AUDIO_URL - 1].toString().trim() : ""
    });
  }

  if (scenes.length === 0) {
    ui.alert(
      "⚠️ No Scene Voiceovers Found",
      "No scenes have VOICEOVER_SYNC text (column 21) for: " + idea.id + "\n\n" +
      "Run Stage 4B (Director Review) first — it fills VOICEOVER_SYNC from the script.",
      ui.ButtonSet.OK
    );
    return;
  }

  // ── Check for already-generated scenes ────────────────────────────────────
  const alreadyDone = scenes.filter(s => s.existingUrl.startsWith("http"));
  if (alreadyDone.length > 0) {
    const regenerate = ui.alert(
      "Scene Audio Already Exists",
      alreadyDone.length + " scene(s) already have audio generated.\n\n" +
      "Re-generate ALL " + scenes.length + " scenes? (Yes)\n" +
      "Or skip already-done scenes and generate only new ones? (No)",
      ui.ButtonSet.YES_NO
    );
    if (regenerate !== ui.Button.YES) {
      // filter to only scenes without audio
      const pending = scenes.filter(s => !s.existingUrl.startsWith("http"));
      if (pending.length === 0) {
        ui.alert("All scenes already have voiceover audio. Nothing to generate.");
        return;
      }
      scenes.length = 0;
      pending.forEach(s => scenes.push(s));
    }
  }

  const confirm = ui.alert(
    "🎙️ Stage 7B — Generate Scene Voiceovers",
    "Generating individual ElevenLabs audio for " + scenes.length + " scene(s):\n\n" +
    scenes.map(s => "  • Scene " + s.sceneNum + " (" + s.sceneType + ") — " +
      s.voiceoverSync.substring(0, 60) + (s.voiceoverSync.length > 60 ? "…" : "")).join("\n") +
    "\n\nLanguage : " + (idea.language || "English") +
    "\nVoice    : " + voiceResult.key +
    "\n\nEach scene generates a separate MP3 saved to Drive.\n" +
    "This may take 1–3 minutes total.\n\nProceed?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const scenesFolder = getOrCreateScenesFolder(idea.id, idea.company);

  let succeeded = 0;
  let failed    = 0;

  scenes.forEach(scene => {
    try {
      Logger.log("Stage 7B: Generating audio for scene " + scene.sceneNum + " (" +
        scene.voiceoverSync.length + " chars)");

      const ttsPayload = {
        text          : speakNumbers_(scene.voiceoverSync),
        model_id      : ELEVENLABS_MODEL,
        voice_settings: {
          stability        : 0.5,
          similarity_boost : 0.75,
          style            : 0.3,
          use_speaker_boost: true
        }
      };

      const response = UrlFetchApp.fetch(
        ELEVENLABS_API_URL + "/" + voiceId + "?output_format=mp3_44100_128",
        {
          method            : "post",
          contentType       : "application/json",
          headers           : { "xi-api-key": apiKey },
          payload           : JSON.stringify(ttsPayload),
          muteHttpExceptions: true
        }
      );

      const code = response.getResponseCode();
      if (code !== 200) {
        throw new Error("ElevenLabs API error " + code + ": " + response.getContentText().substring(0, 200));
      }

      const fileName  = idea.id + "_scene_" + scene.sceneNum + "_voiceover.mp3";
      const audioBlob = response.getBlob().setContentType("audio/mpeg").setName(fileName);

      // Delete any existing file with the same name
      const existing = scenesFolder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);

      const audioFile = scenesFolder.createFile(audioBlob);
      const audioUrl  = audioFile.getUrl();

      // Write URL back to Visual Library col 23
      visualSheet.getRange(scene.row, COL_VISUAL_EXTENDED.VOICEOVER_AUDIO_URL).setValue(audioUrl);
      SpreadsheetApp.flush();

      succeeded++;
      Logger.log("Stage 7B: ✅ Scene " + scene.sceneNum + " → " + fileName);

      // Rate-limit pause: ElevenLabs allows ~10 concurrent, but GAS is sequential
      // 300ms between calls avoids 429 rate-limit errors on long batches
      Utilities.sleep(300);

    } catch (err) {
      failed++;
      Logger.log("Stage 7B: ❌ Scene " + scene.sceneNum + " failed: " + err.message);
      logError("Stage 7B — Scene Voiceover", idea.id, "Scene " + scene.sceneNum, err.message);
    }
  });

  updatePipelineStatus_(idea.id, "S7B", failed === 0 ? "✅" : (succeeded > 0 ? "⚠️" : "❌"));

  ui.alert(
    succeeded === scenes.length ? "✅ Stage 7B Complete — Scene Voiceovers Generated" : "⚠️ Stage 7B Partial",
    "Generated: " + succeeded + "/" + scenes.length + " scene audio files\n" +
    (failed > 0 ? "Failed : " + failed + " (see Error Log)\n\n" : "\n") +
    "Audio files saved to:\n📁 " + idea.company + " / Scenes /\n\n" +
    "URLs written to Visual Library column 23 (VOICEOVER_AUDIO_URL).\n" +
    "Shotstack (Stage 9B) will use these for precise scene timing.",
    ui.ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 8A — Process Scenes (Dual Path: Pexels + KlingAI)
// Pexels scenes → search API → write best video URL to Link column immediately
// KlingAI scenes → submit job → store Task ID → run 8B later to collect
// ════════════════════════════════════════════════════════════════════════════════
function submitVideoClips() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  const props      = PropertiesService.getScriptProperties();
  const pexelsKey  = props.getProperty("PEXELS_API_KEY");
  const accessKey  = props.getProperty("KLING_ACCESS_KEY");
  const secretKey  = props.getProperty("KLING_SECRET_KEY");

  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const status = visualData[i][COL_VISUAL.STATUS - 1].toString().trim();
    if (status === "Ready" || status === "Done" || status === "Skip") continue;
    scenes.push({
      row         : i + 1,
      num         : visualData[i][COL_VISUAL.SCENE_NUM   - 1],
      type        : visualData[i][COL_VISUAL.SCENE_TYPE  - 1].toString().trim(),
      description : visualData[i][COL_VISUAL.DESCRIPTION - 1].toString().trim(),
      source      : visualData[i][COL_VISUAL.SOURCE      - 1].toString().trim(),
      link        : visualData[i][COL_VISUAL.LINK        - 1].toString().trim(),
      status      : status
    });
  }

  if (scenes.length === 0) {
    ui.alert("No pending scenes found for: " + idea.id); return;
  }

  const pexelsScenes = scenes.filter(s => s.source === "Pexels" || s.source === "Getty");
  const klingScenes  = scenes.filter(s => s.source === "KlingAI").slice(0, idea.maxAiClips || 5);
  const otherScenes  = scenes.filter(s =>
    s.source !== "Pexels" && s.source !== "Getty" &&
    s.source !== "KlingAI"
  ); // Includes CapCut, Canva, Original — Checkpoint and Timeline scenes all use CapCut

  const skippedScenes = scenes.filter(s => s.status === "Skip").length;

  const confirm = ui.alert(
    "Stage 8A — Process Scenes",
    "Processing scenes for: " + idea.company + "\n" +
    "Production Mode: " + idea.productionMode + "\n" +
    "Max AI Clips: " + idea.maxAiClips + "\n\n" +
    "📷 Pexels search: " + pexelsScenes.length + " scenes (instant)\n" +
    "🎬 KlingAI generate: " + klingScenes.length + " scenes (async — run 8B later)\n" +
    "🎨 Manual (CapCut/Canva/Other): " + otherScenes.length + " scenes (see Assembly Guide)\n" +
    (skippedScenes > 0 ? "⏭️ Skipped: " + skippedScenes + " scenes (Status = Skip)\n" : "") +
    "\nProceed?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  let pexelsSuccess = 0, pexelsFail = 0;
  let klingSuccess  = 0, klingFail  = 0;

  // ── PATH 1: PEXELS ────────────────────────────────────────────────────────
  if (pexelsScenes.length > 0 && !pexelsKey) {
    ui.alert(
      "⚠️ Pexels API Key Missing",
      "PEXELS_API_KEY not found in Script Properties.\n\n" +
      "Get a free key at pexels.com/api then add it to Script Properties.\n" +
      "Skipping " + pexelsScenes.length + " Pexels scenes for now.",
      ui.ButtonSet.OK
    );
  } else {
    pexelsScenes.forEach(scene => {
      try {
        const rawLink    = scene.link;
        const queryMatch = rawLink.match(/Search:\s*(.+)/i);
        const query      = queryMatch
          ? queryMatch[1].trim()
          : scene.description.substring(0, 50);

        const response = UrlFetchApp.fetch(
          PEXELS_API_URL + "?query=" + encodeURIComponent(query) +
          "&per_page=5&orientation=landscape",
          {
            headers           : { "Authorization": pexelsKey },
            muteHttpExceptions: true
          }
        );

        const code = response.getResponseCode();
        const body = JSON.parse(response.getContentText());

        if (code !== 200 || !body.videos || body.videos.length === 0) {
          throw new Error("No results for query: " + query);
        }

        const video   = body.videos[0];
        const hdFile  = video.video_files.find(f => f.quality === "hd") ||
                        video.video_files[0];
        const directUrl = hdFile ? hdFile.link : video.url;

        visualSheet.getRange(scene.row, COL_VISUAL.LINK  ).clearDataValidations();
        visualSheet.getRange(scene.row, COL_VISUAL.LINK  ).setValue(directUrl);
        visualSheet.getRange(scene.row, COL_VISUAL.STATUS).setValue("Ready");
        pexelsSuccess++;

        Logger.log("Scene " + scene.num + " Pexels: " + directUrl);
        Utilities.sleep(500);

      } catch (err) {
        pexelsFail++;
        logError("Stage 8A — Scene " + scene.num, idea.id, "Pexels Error", err.message);
      }
    });
  }

  // ── PATH 2: KLING AI ──────────────────────────────────────────────────────
  if (klingScenes.length > 0 && (!accessKey || !secretKey)) {
    ui.alert(
      "⚠️ KlingAI Keys Missing",
      "KLING_ACCESS_KEY or KLING_SECRET_KEY not found.\n" +
      "Skipping " + klingScenes.length + " KlingAI scenes.",
      ui.ButtonSet.OK
    );
  klingScenes.forEach(scene => {
      try {
        const token      = generateKlingJWT(accessKey, secretKey);
        const rawLink    = scene.link;
        const aiMatch    = rawLink.match(/AI:\s*(.+)/i);
        const aiPrompt   = aiMatch ? aiMatch[1].trim() : scene.description;

        const submitPayload = {
          model          : KLING_MODEL,
          prompt         : "Cinematic documentary B-roll. " + aiPrompt +
                           ". Dark professional tone, sharp cinematic focus. " +
                           "CRITICAL: NO text NO words NO letters NO numbers NO signs " +
                           "NO captions NO watermarks anywhere in frame. Pure visual only.",
          negative_prompt: "text, letters, words, numbers, captions, subtitles, watermark, " +
                           "logo, signs, writing, typography, fonts, labels, titles, " +
                           "blurry, low quality, amateur, distorted",
          cfg_scale      : 0.5,
          mode           : "std",
          duration       : "5"
        };

        const response = UrlFetchApp.fetch(KLING_API_URL, {
          method            : "post",
          contentType       : "application/json",
          headers           : { "Authorization": "Bearer " + token },
          payload           : JSON.stringify(submitPayload),
          muteHttpExceptions: true
        });

        const code = response.getResponseCode();
        const body = JSON.parse(response.getContentText());

        if (code !== 200 || body.code !== 0) {
          throw new Error("Submit failed: " + JSON.stringify(body));
        }

        const taskId = body.data.task_id;
        visualSheet.getRange(scene.row, COL_VISUAL.TASK_ID).setValue(taskId);
        visualSheet.getRange(scene.row, COL_VISUAL.STATUS ).setValue("Submitted");
        klingSuccess++;

        Logger.log("Scene " + scene.num + " KlingAI submitted. Task ID: " + taskId);
        Utilities.sleep(2000);

      } catch (err) {
        klingFail++;
        logError("Stage 8A — Scene " + scene.num, idea.id, "KlingAI Submit Error", err.message);
      }
    });
  }

  // ── PATH 3: OTHER (Canva, Original) — skip, manual creation needed ────────
  otherScenes.forEach(scene => {
    Logger.log("Scene " + scene.num + " skipped (" + scene.source + ") — manual");
  });

  let summary = "Stage 8A Results\n\n";
  summary    += "📷 Pexels: "  + pexelsSuccess + " found" +
                (pexelsFail > 0 ? ", " + pexelsFail + " failed" : "") + "\n";
  summary    += "🎬 KlingAI: " + klingSuccess  + " submitted" +
                (klingFail > 0 ? ", " + klingFail + " failed" : "")  + "\n";
  summary    += "📋 Manual: "  + otherScenes.length + " need Canva/manual creation\n";

  if (klingSuccess > 0) {
    summary += "\n⏳ Wait 5–10 minutes then run Stage 8B to collect KlingAI clips.";
  }

  ui.alert("✅ Stage 8A Complete", summary, ui.ButtonSet.OK);
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 8B — Collect Video Clips (KlingAI)
// Reads stored Task IDs, polls KlingAI, writes completed clip URLs back to sheet
// ════════════════════════════════════════════════════════════════════════════════
function collectVideoClips() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  const props     = PropertiesService.getScriptProperties();
  const accessKey = props.getProperty("KLING_ACCESS_KEY");
  const secretKey = props.getProperty("KLING_SECRET_KEY");

  if (!accessKey) { ui.alert("KLING_ACCESS_KEY missing from Script Properties."); return; }
  if (!secretKey) { ui.alert("KLING_SECRET_KEY missing from Script Properties."); return; }

  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const submitted   = [];

  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const status = visualData[i][COL_VISUAL.STATUS  - 1].toString().trim();
    const taskId = visualData[i][COL_VISUAL.TASK_ID - 1].toString().trim();
    if (status === "Submitted" && taskId !== "") {
      submitted.push({
        row    : i + 1,
        num    : visualData[i][COL_VISUAL.SCENE_NUM - 1],
        taskId : taskId
      });
    }
  }

  if (submitted.length === 0) {
    ui.alert(
      "No submitted scenes found",
      "No scenes with status 'Submitted' found for: " + idea.id +
      "\nRun Stage 8A first.",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    "Stage 8B — Collecting Video Clips",
    "Checking status of " + submitted.length + " submitted scenes for: " + idea.company +
    "\n\nThis will take about 30–60 seconds.",
    ui.ButtonSet.OK
  );

  let ready   = 0;
  let pending = 0;
  let failed  = 0;
  const scenesFolder    = getOrCreateScenesFolder(idea.id, idea.company);
  const scenesFolderUrl = "https://drive.google.com/drive/folders/" + scenesFolder.getId();

  submitted.forEach(scene => {
    try {
      const token = generateKlingJWT(accessKey, secretKey);

      const statusResponse = UrlFetchApp.fetch(
        KLING_API_STATUS_URL + "/" + scene.taskId,
        {
          method            : "get",
          headers           : { "Authorization": "Bearer " + token },
          muteHttpExceptions: true
        }
      );

      const statusBody   = JSON.parse(statusResponse.getContentText());
      const taskStatus   = statusBody.data && statusBody.data.task_status;

      Logger.log("Scene " + scene.num + " status: " + taskStatus);

      if (taskStatus === "succeed") {
        const klingUrl = statusBody.data.task_result &&
                         statusBody.data.task_result.videos &&
                         statusBody.data.task_result.videos[0] &&
                         statusBody.data.task_result.videos[0].url;

        if (klingUrl) {
          // ── Download clip and save to Drive Scenes folder ───────────────
          try {
            const scenesFolder = getOrCreateScenesFolder(idea.id, idea.company);
            const fileName     = idea.id + " — Scene-" +
                                 String(scene.num).padStart(2, "0") + ".mp4";

            // Delete existing clip if any
            const existing = scenesFolder.getFilesByName(fileName);
            while (existing.hasNext()) existing.next().setTrashed(true);

            // Download from KlingAI and save to Drive
            const videoBlob = UrlFetchApp.fetch(klingUrl, {
              muteHttpExceptions: true
            }).getBlob().setContentType("video/mp4");

            const savedFile = scenesFolder.createFile(
              videoBlob.setName(fileName)
            );
            const driveUrl = "https://drive.google.com/file/d/" +
                             savedFile.getId() + "/view";

            // Write Drive link to AI_CLIP_URL — never touch LINK column
            visualSheet.getRange(scene.row, COL_VISUAL.AI_CLIP_URL).setValue(driveUrl);
            visualSheet.getRange(scene.row, COL_VISUAL.STATUS      ).setValue("Ready");
            ready++;

            Logger.log("Scene " + scene.num + " saved to Drive: " + driveUrl);

          } catch (downloadErr) {
            // If download fails, at least save the temporary KlingAI URL
            visualSheet.getRange(scene.row, COL_VISUAL.AI_CLIP_URL).setValue(klingUrl);
            visualSheet.getRange(scene.row, COL_VISUAL.STATUS      ).setValue("Ready");
            ready++;
            Logger.log("Scene " + scene.num + " Drive save failed, using KlingAI URL: " + downloadErr.message);
          }
        }

      } else if (taskStatus === "failed") {
        visualSheet.getRange(scene.row, COL_VISUAL.STATUS).setValue("Error");
        logError("Stage 8B — Scene " + scene.num, idea.id, "KlingAI Processing Failed",
          "Task " + scene.taskId + " failed on KlingAI side");
        failed++;

      } else {
        // Still processing
        pending++;
        Logger.log("Scene " + scene.num + " still processing: " + taskStatus);
      }

    } catch (err) {
      failed++;
      logError("Stage 8B — Scene " + scene.num, idea.id, "KlingAI Collect Error", err.message);
    }
  });

  // Write Scenes Folder link to Publishing Tracker if any clips are ready
  if (ready > 0) {
    writePublishingLink(idea.id, idea.company, master,
      COL_PUBLISHING.SCENES_FOLDER, scenesFolderUrl, "📁 Open Scenes");
  }

  let message = "Results for: " + idea.id + "\n\n";
  message    += "✅ Ready: "      + ready   + " clips\n";
  message    += "⏳ Still processing: " + pending + " clips\n";
  message    += (failed > 0 ? "❌ Failed: " + failed + " clips (see Error Log)\n" : "");

  if (pending > 0) {
    message += "\nSome clips are still processing.\n" +
               "Wait a few more minutes and run Stage 8B again.";
  } else if (ready > 0 && failed === 0) {
    message += "\nAll clips collected successfully!\n" +
               "Check Visual Library → Link column for clip URLs.";
  }

  ui.alert(
    pending > 0 ? "⏳ Stage 8B — Partial Results" : "✅ Stage 8B Complete",
    message,
    ui.ButtonSet.OK
  );
}

// ── Build a scene-specific video prompt for KlingAI ──────────────────────────
function buildScenePrompt(scene, idea, master) {
  return "Cinematic documentary B-roll footage. " +
    "Topic: " + idea.company + ". " +
    "Scene: " + scene.description + ". " +
    "Style: dark professional tone, sharp cinematic focus, high production quality. " +
    "CRITICAL: NO text, NO words, NO letters, NO numbers, NO signs, NO captions, " +
    "NO watermarks, NO logos, NO subtitles anywhere in the frame. " +
    "Pure visual footage only. No text of any kind. " +
    "Duration: 5 seconds.";
}

// ── Generate KlingAI JWT authentication token ─────────────────────────────────
function generateKlingJWT(accessKey, secretKey) {
  const header  = { alg: "HS256", typ: "JWT" };
  const now     = Math.floor(Date.now() / 1000);
  const payload = {
    iss : accessKey,
    exp : now + 1800,  // 30 min expiry
    nbf : now - 5
  };

  const base64Header  = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, "");
  const base64Payload = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=+$/, "");
  const signingInput  = base64Header + "." + base64Payload;

  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(signingInput, secretKey)
  ).replace(/=+$/, "");

  return signingInput + "." + signature;
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 9 — Build Assembly Package (CapCut)
// Creates a complete production-ready assembly guide as a Google Doc
// References the GovernX CapCut template with exact instructions per scene
// ════════════════════════════════════════════════════════════════════════════════
function buildAssemblyPackage() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Pull script data ──────────────────────────────────────────────────────
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let   script      = null;

  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      script = {
        hook      : scriptData[i][COL_SCRIPT.HOOK             - 1],
        voiceover : scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1],
        closing   : scriptData[i][COL_SCRIPT.GRC_BPR_CLOSING  - 1],
        format    : scriptData[i][COL_SCRIPT.TARGET_FORMAT    - 1],
        cta       : scriptData[i][COL_SCRIPT.CALL_TO_ACTION   - 1],
        sections  : scriptData[i][COL_SCRIPT.SECTIONS         - 1]
      };
      break;
    }
  }

  if (!script) {
    ui.alert("No script found for: " + idea.id + "\nRun Stage 3 first.");
    return;
  }

  // ── Pull scene data ───────────────────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    scenes.push({
      num             : visualData[i][COL_VISUAL.SCENE_NUM        - 1],
      type            : visualData[i][COL_VISUAL.SCENE_TYPE       - 1],
      description     : visualData[i][COL_VISUAL.DESCRIPTION      - 1],
      source          : visualData[i][COL_VISUAL.SOURCE           - 1],
      link            : visualData[i][COL_VISUAL.LINK             - 1],
      timestamp       : visualData[i][COL_VISUAL.TIMESTAMP        - 1],
      license         : visualData[i][COL_VISUAL.LICENSE          - 1],
      status          : visualData[i][COL_VISUAL.STATUS           - 1],
      aiClipUrl       : visualData[i][COL_VISUAL.AI_CLIP_URL      - 1],
      checkpointDate  : visualData[i][COL_VISUAL.CHECKPOINT_DATE  - 1],
      checkpointEvent : visualData[i][COL_VISUAL.CHECKPOINT_EVENT - 1],
      checkpointAngle : visualData[i][COL_VISUAL.CHECKPOINT_ANGLE - 1]
    });
  }

  if (scenes.length === 0) {
    ui.alert("No scenes found for: " + idea.id + "\nRun Stage 4 first.");
    return;
  }

  // ── Pull voiceover audio link ─────────────────────────────────────────────
  const pubSheet  = ss.getSheetByName(SHEET.PUBLISHING);
  let   audioLink = "";
  if (pubSheet) {
    const pubData = pubSheet.getDataRange().getValues();
    for (let i = 1; i < pubData.length; i++) {
      if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === idea.id) {
        audioLink = pubData[i][COL_PUBLISHING.VOICEOVER_AUDIO - 1] || "";
        break;
      }
    }
  }

  const isArabic  = idea.language === "Arabic" || idea.language === "Bilingual";
  const divider   = "═".repeat(60);
  const line      = "─".repeat(60);

  // ── GovernX Template Reference ───────────────────────────────────────────
  const TEMPLATE = {
    canvas     : "1920×1080 | 24fps | Background: #0A0A0A",
    titleFont  : isArabic ? "Cairo Black" : "Montserrat Black",
    bodyFont   : isArabic ? "Cairo Bold" : "Montserrat Bold",
    subFont    : isArabic ? "Cairo Regular" : "Montserrat Regular",
    align      : isArabic ? "Right (RTL)" : "Left",
    direction  : isArabic ? "⚠️ Enable RTL in CapCut: Text → Advanced → Right to Left" : "Left to Right",
    accent     : "#FF0000",
    textPrimary: "#FFFFFF",
    textMuted  : "#CCCCCC",
    overlay    : "#000000 at 60% opacity",
    styles: {
      IMPACT_STAT    : (isArabic ? "Cairo Black" : "Montserrat Black") + " | 120pt | #FFFFFF | Center | Fade up 0.3s | Red underline #FF0000",
      CHAPTER_TITLE  : (isArabic ? "Cairo Black" : "Montserrat Bold") + " | 52pt | #FFFFFF | Lower third " + (isArabic ? "RIGHT (RTL)" : "LEFT") + " | Slide " + (isArabic ? "left" : "right") + " 0.4s | Red bar #FF0000",
      CONTEXT_LABEL  : (isArabic ? "Cairo Regular" : "Montserrat Regular") + " | 28pt | #CCCCCC | Bottom " + (isArabic ? "right" : "left") + " 80px | Fade 0.2s | Black pill bg",
      GRC_CLOSING    : (isArabic ? "Cairo Bold" : "Montserrat Bold Italic") + " | 38pt | #FF0000 | Center | Fade slow 0.8s | Full black bg",
      LOWER_THIRD    : (isArabic ? "Cairo Regular" : "Montserrat Regular") + " | 22pt | #999999 | Bottom " + (isArabic ? "right" : "left") + " | No animation"
    },
    transitions: {
      broll    : "Cross dissolve | 0.5s",
      toText   : "Fade to black | 0.3s",
      fromText : "Fade from black | 0.3s",
      impact   : "Hard cut | No transition"
    },
    audio: {
      voiceover : "0dB | No effects",
      music     : "-18dB background",
      sfx       : "-12dB"
    }
  };

  // ── Build Assembly Guide document ─────────────────────────────────────────
  let doc = "";

  // ── HEADER ────────────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "GOVERNX — CAPCUT ASSEMBLY GUIDE\n";
  doc += divider + "\n";
  doc += "Content ID  : " + idea.id + "\n";
  doc += "Title       : " + (master ? master.title : idea.company) + "\n";
  doc += "Format      : " + script.format + "\n";
  doc += "Language    : " + idea.language + "\n";
  doc += "Scenes      : " + scenes.length + "\n";
  doc += "Generated   : " + new Date().toLocaleString() + "\n";
  doc += divider + "\n\n";

  // ── SECTION 1: PROJECT SETUP ──────────────────────────────────────────────
  doc += "SECTION 1 — CAPCUT PROJECT SETUP\n";
  doc += line + "\n\n";
  doc += "STEP 1 — CREATE NEW PROJECT\n";
  doc += "• Open CapCut Desktop\n";
  doc += "• Click 'New Project'\n";
  doc += "• Resolution: 1920×1080\n";
  doc += "• Frame Rate: 24fps\n";
  doc += "• Click 'Create'\n";
  if (isArabic) {
    doc += "\n⚠️ ARABIC VIDEO — TEXT DIRECTION:\n";
    doc += "• For all Arabic text layers: Text → Advanced → enable 'Right to Left'\n";
    doc += "• Font to install: Cairo (download free from Google Fonts)\n";
    doc += "• Text alignment: Right-aligned for all overlays\n";
    doc += "• Checkpoint cards: Date and Event text must be RTL enabled\n\n";
  }

  doc += "STEP 2 — IMPORT ASSETS\n";
  doc += "• Click 'Import' in the media panel\n";
  doc += "• Import all scene clips from your Drive Scenes folder\n";
  doc += "• Import voiceover audio: " + (audioLink || "See Publishing Tracker → Voiceover Audio") + "\n";
  doc += "• Clips will appear in your media library\n\n";

  doc += "STEP 3 — SET BACKGROUND\n";
  doc += "• Add a solid color background track\n";
  doc += "• Color: #0A0A0A (near black)\n";
  doc += "• Extend it to full video duration\n\n";

  doc += "STEP 4 — ADD VOICEOVER\n";
  doc += "• Drag voiceover MP3 to Audio track\n";
  doc += "• Set volume to 0dB\n";
  doc += "• This is your timing anchor — all visuals sync to voice\n\n";

  // ── SECTION 2: TIMELINE ASSEMBLY ─────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 2 — TIMELINE ASSEMBLY (Scene by Scene)\n";
  doc += line + "\n\n";
  doc += "⚠️  KEY RULE: Always place clips on Video Track 1.\n";
  doc += "     Text overlays go on Video Track 2 (above clips).\n";
  doc += "     Voiceover stays on Audio Track 1.\n\n";

  scenes.forEach((scene, index) => {
    const sceneNum    = String(scene.num).padStart(2, "0");
    const nextScene   = scenes[index + 1];
    const duration    = calculateDuration(scene.timestamp, nextScene ? nextScene.timestamp : null, script.format);
    const transition  = getTransition(scene.type, TEMPLATE);

    doc += "────────────────────────────────────────\n";
    doc += "SCENE " + sceneNum + " | " + scene.timestamp + " | " + scene.type.toUpperCase() + "\n";
    doc += "────────────────────────────────────────\n";
    doc += "Duration     : ~" + duration + " seconds\n";
    doc += "Source       : " + scene.source + "\n";

    if (scene.status === "Ready" && scene.link) {
      doc += "Clip URL     : " + scene.link + "\n";
    } else if (scene.aiClipUrl) {
      doc += "AI Clip URL  : " + scene.aiClipUrl + "\n";
    } else {
      doc += "Clip Status  : " + (scene.status || "Needed") + " — source: " + scene.link + "\n";
    }

    doc += "Description  : " + scene.description + "\n\n";

    // ── CapCut instructions by scene type ──────────────────────────────────
    if (scene.type === "Stock" || scene.type === "AI Generated") {
      doc += "CAPCUT STEPS:\n";
      doc += "1. Drag clip to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Trim to ~" + duration + " seconds\n";
      doc += "3. Add transition: " + transition + "\n";
      if (scene.license === "Fair Use") {
        doc += "4. Add LOWER THIRD attribution (Style 5):\n";
        doc += "   • Text: 'Source: " + scene.source + "'\n";
        doc += "   • " + TEMPLATE.styles.LOWER_THIRD + "\n";
        doc += "   • Duration: 3 seconds from clip start\n";
      }

    } else if (scene.type === "Text" && scene.source === "CapCut") {
      doc += "CAPCUT STEPS (build directly in CapCut):\n";
      doc += "1. Add black solid color to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Duration: ~" + duration + " seconds\n";
      doc += "3. Add text overlay on Video Track 2 — Style 2 (CHAPTER TITLE):\n";
      doc += "   • Content: " + scene.description + "\n";
      doc += "   • " + TEMPLATE.styles.CHAPTER_TITLE + "\n";
      doc += "   • Align: " + TEMPLATE.align + "\n";
      doc += "4. For sequential stat punch-ins:\n";
      doc += "   • Add each stat as a SEPARATE text layer\n";
      doc += "   • Stagger start times: 0.0s, 0.8s, 1.6s, 2.4s\n";
      doc += "   • Animation: Pop or Bounce in | 0.2s each\n";
      doc += "   • Numbers/stats: Style 1 (IMPACT STAT) | #FF0000\n";
      doc += "   • Labels: Style 2 (CHAPTER TITLE) | #FFFFFF\n";
      doc += "5. Transition in: " + TEMPLATE.transitions.fromText + "\n";
      doc += "5. Transition out: " + TEMPLATE.transitions.toText + "\n";

    } else if (scene.type === "Text") {
      doc += "CAPCUT STEPS:\n";
      doc += "1. Add black solid color to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Duration: ~" + duration + " seconds\n";
      doc += "3. Add text overlay on Video Track 2:\n";
      doc += "   • Content: " + scene.description + "\n";
      doc += "   • Style: " + TEMPLATE.styles.CHAPTER_TITLE + "\n";
      doc += "   • Align: " + TEMPLATE.align + "\n";
      doc += "4. Transition in: " + TEMPLATE.transitions.fromText + "\n";
      doc += "5. Transition out: " + TEMPLATE.transitions.toText + "\n";

    } else if (scene.type === "Animation") {
      doc += "CAPCUT STEPS:\n";
      doc += "1. Create this graphic in Canva first:\n";
      doc += "   • " + scene.description + "\n";
      doc += "   • Canvas: 1920×1080 | Background: #0A0A0A\n";
      doc += "   • Export as MP4 (animated) or PNG (static)\n";
      doc += "2. Import into CapCut → drag to Video Track 1 at " + scene.timestamp + "\n";
      doc += "3. Duration: ~" + duration + " seconds\n";
      doc += "4. Add IMPACT STAT text overlay if needed:\n";
      doc += "   • Style: " + TEMPLATE.styles.IMPACT_STAT + "\n";

    } else if (scene.type === "Checkpoint") {
      const cpDate  = scene.checkpointDate  || extractCheckpointPart(scene.description, "DATE");
      const cpEvent = scene.checkpointEvent || extractCheckpointPart(scene.description, "EVENT");
      const cpAngle = scene.checkpointAngle || extractCheckpointPart(scene.description, "ANGLE");

      doc += "CAPCUT STEPS — CHECKPOINT CARD (GovernX Signature):\n";
      doc += "1. Add black solid (#0A0A0A) to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Duration: 3 seconds\n";
      doc += "3. Transition in: Hard cut | Transition out: Fade to black 0.3s\n\n";
      doc += "4. TEXT LAYERS (build all on Video Track 2):\n\n";
      doc += "   Layer 1 — RED LINE (appears at 0.0s):\n";
      doc += "   • Shape: Rectangle 400px × 3px | Color: #FF0000\n";
      doc += "   • Position: Center screen, upper third\n";
      doc += "   • Animation: Scale from 0 → full width | 0.3s\n\n";
      doc += "   Layer 2 — DATE (appears at 0.3s):\n";
      doc += "   • Text: " + cpDate + "\n";
      doc += "   • Font: Montserrat Black | 72pt | #FFFFFF | Center\n";
      doc += "   • Animation: Pop (scale up) | 0.2s\n\n";
      doc += "   Layer 3 — RED DIVIDER (appears at 0.8s):\n";
      doc += "   • Shape: Rectangle 300px × 3px | Color: #FF0000\n";
      doc += "   • Position: Center, below date\n";
      doc += "   • Animation: Slide from left | 0.3s\n\n";
      doc += "   Layer 4 — EVENT (appears at 1.2s):\n";
      doc += "   • Text: " + cpEvent + "\n";
      doc += "   • Font: Montserrat Bold | 42pt | #FFFFFF | Center\n";
      doc += "   • Animation: Fade up | 0.3s\n\n";
      doc += "   Layer 5 — ANGLE TAG (appears at 1.8s):\n";
      doc += "   • Text: " + cpAngle.toUpperCase() + "\n";
      doc += "   • Font: Montserrat Bold | 28pt | #FF0000 | Center\n";
      doc += "   • Animation: Pop | 0.2s\n\n";
      doc += "⚠️ Add bass impact SFX at 0.3s and 1.8s (−12dB)\n";

    } else if (scene.type === "Timeline") {
      const checkpointScenes = scenes.filter(s => s.type === "Checkpoint");
      const cpCount = checkpointScenes.length;
      const totalDur = (cpCount * 0.8 + 3).toFixed(1);

      doc += "CAPCUT STEPS — FINAL TIMELINE REVEAL (GovernX Signature Ending):\n";
      doc += "1. Add black solid (#0A0A0A) to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Duration: " + totalDur + " seconds (" + cpCount + " checkpoints × 0.8s + 3s hold)\n";
      doc += "3. Transition in: Hard cut | Transition out: Fade to black 0.8s\n\n";
      doc += "4. LAYER 1 — HEADER (appears at 0.0s):\n";
      doc += "   • Text: THE COLLAPSE MAP\n";
      doc += "   • Font: Montserrat Black | 28pt | #FF0000 | Center top\n";
      doc += "   • Animation: Fade in | 0.3s\n\n";
      doc += "5. CHECKPOINT LINES — bottom to top reveal:\n\n";

      // Reverse for bottom-to-top (root cause first, outcome last)
      const reversedCPs = [...checkpointScenes].reverse();
      reversedCPs.forEach((cp, idx) => {
        const delay    = (0.5 + idx * 0.8).toFixed(1);
        const cpDate   = cp.checkpointDate  || "—";
        const cpEvent  = cp.checkpointEvent || cp.description.substring(0, 50);
        const cpAngle  = cp.checkpointAngle || "";
        const isRoot   = idx === 0;
        const isOutcome = idx === reversedCPs.length - 1;

        doc += "   Layer " + (idx + 2) + " — appears at " + delay + "s";
        if (isRoot)    doc += " [ROOT CAUSE]";
        if (isOutcome) doc += " [OUTCOME]";
        doc += ":\n";
        doc += "   • Date: " + cpDate + " | Event: " + cpEvent + "\n";
        if (cpAngle) doc += "   • Angle: " + cpAngle + "\n";
        doc += "   • Font: Montserrat Regular | 22pt | " + (isRoot ? "#FF0000 Bold" : "#FFFFFF") + " | Center\n";
        doc += "   • Animation: Slide up from bottom | 0.3s\n";
        if (idx < reversedCPs.length - 1) {
          doc += "   • Add ↑ arrow below | #FF0000 | 20pt\n";
        }
        doc += "\n";
      });

      doc += "6. ROOT CAUSE line (appears at " + (0.5 + cpCount * 0.8).toFixed(1) + "s):\n";
      doc += "   • Text: ROOT: " + (master ? master.discipline : "GOVERNANCE FAILURE") + "\n";
      doc += "   • Font: Montserrat Black | 26pt | #FF0000 | Center\n";
      doc += "   • Animation: Pop | 0.3s | Add red underline below\n\n";
      doc += "7. HOLD: all checkpoints visible for 2 seconds\n";
      doc += "8. Add deep cinematic bass note at hold moment (−10dB)\n";

    } else if (scene.type === "Minimal") {
      doc += "CAPCUT STEPS:\n";
      doc += "1. Add black solid color to Video Track 1 at " + scene.timestamp + "\n";
      doc += "2. Duration: ~" + duration + " seconds\n";
      doc += "3. No text overlay — voiceover only\n";
      doc += "4. Transition: " + TEMPLATE.transitions.impact + "\n";
    }

    doc += "\n";
  });

  // ── SECTION 3: TEXT OVERLAYS SUMMARY ─────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 3 — TEXT OVERLAY QUICK REFERENCE\n";
  doc += line + "\n\n";
  doc += "Use these exact styles for all text in this video:\n\n";
  doc += "STYLE 1 — IMPACT STAT (key numbers)\n";
  doc += "  " + TEMPLATE.styles.IMPACT_STAT + "\n\n";
  doc += "STYLE 2 — CHAPTER TITLE (section headers)\n";
  doc += "  " + TEMPLATE.styles.CHAPTER_TITLE + "\n\n";
  doc += "STYLE 3 — CONTEXT LABEL (names, dates, roles)\n";
  doc += "  " + TEMPLATE.styles.CONTEXT_LABEL + "\n\n";
  doc += "STYLE 4 — GRC CLOSING (GovernX signature ending)\n";
  doc += "  " + TEMPLATE.styles.GRC_CLOSING + "\n\n";
  doc += "STYLE 5 — LOWER THIRD (source attribution)\n";
  doc += "  " + TEMPLATE.styles.LOWER_THIRD + "\n\n";

  // ── SECTION 4: GRC CLOSING ARGUMENT ──────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 4 — GRC CLOSING ARGUMENT CARD\n";
  doc += line + "\n\n";
  doc += "Add this as the FINAL text card before the CTA:\n\n";
  doc += "TEXT CONTENT:\n";
  doc += script.closing + "\n\n";
  doc += "CAPCUT STEPS:\n";
  doc += "1. Add 4-second black solid at end of timeline\n";
  doc += "2. Add text overlay — Style 4 (GRC CLOSING):\n";
  doc += "   • " + TEMPLATE.styles.GRC_CLOSING + "\n";
  doc += "   • " + TEMPLATE.align + " aligned\n";
  doc += "3. Transition in: Fade from black 0.8s\n\n";

  // ── SECTION 5: CALL TO ACTION ─────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 5 — CALL TO ACTION CARD\n";
  doc += line + "\n\n";
  doc += "CTA: " + script.cta + "\n\n";
  doc += "CAPCUT STEPS:\n";
  doc += "1. Add 3-second black solid after GRC closing card\n";
  doc += "2. Add text — Style 2 (CHAPTER TITLE):\n";
  doc += "   • " + TEMPLATE.styles.CHAPTER_TITLE + "\n";
  doc += "3. Add GovernX logo/watermark bottom " + (isArabic ? "left" : "right") + "\n\n";

  // ── SECTION 6: AUDIO SETTINGS ────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 6 — AUDIO SETTINGS\n";
  doc += line + "\n\n";
  doc += "VOICEOVER (Track 1): " + TEMPLATE.audio.voiceover + "\n";
  doc += "BACKGROUND MUSIC (Track 2): " + TEMPLATE.audio.music + "\n";
  doc += "  Recommended style: Cinematic minimal, no lyrics, dark ambient\n";
  doc += "  Suggested search: 'dark corporate documentary music no copyright'\n";
  doc += "SOUND EFFECTS (Track 3 if used): " + TEMPLATE.audio.sfx + "\n\n";

  // ── SECTION 7: EXPORT SETTINGS ───────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION 7 — EXPORT SETTINGS FOR YOUTUBE\n";
  doc += line + "\n\n";
  doc += "Resolution   : 1920×1080 (1080p)\n";
  doc += "Frame Rate   : 24fps\n";
  doc += "Format       : MP4\n";
  doc += "Codec        : H.264\n";
  doc += "Bitrate      : High (recommended)\n";
  doc += "Color Space  : SDR\n\n";
  doc += "In CapCut: File → Export → select settings above → Export\n\n";

  doc += divider + "\n";
  doc += "END OF ASSEMBLY GUIDE — " + idea.id + "\n";
  doc += divider + "\n";

  // ── Save to Drive ─────────────────────────────────────────────────────────
  try {
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
    const docTitle      = idea.id + " — " + idea.company + " — Assembly Guide";

    // Delete existing if any
    const existing = contentFolder.getFilesByName(docTitle);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const newDoc = DocumentApp.create(docTitle);
    const body   = newDoc.getBody();

    // Apply basic formatting — monospace for readability
    body.setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: "Courier New" });
    body.setText(doc);
    newDoc.saveAndClose();

    const file = DriveApp.getFileById(newDoc.getId());
    contentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const docUrl = newDoc.getUrl();

    // Write clickable link to Publishing Tracker → Assembly Guide column
    writePublishingLink(idea.id, idea.company, master,
      COL_PUBLISHING.ASSEMBLY_GUIDE, docUrl, "📋 Open Guide");

    ui.alert(
      "✅ Stage 9 Complete — Assembly Guide Created",
      "Saved to Google Drive:\n" +
      "📁 " + DRIVE_FOLDER_NAME + " / " + idea.id + " — " + idea.company + "\n" +
      "📋 " + docTitle + "\n\n" +
      "The guide contains:\n" +
      "• CapCut project setup instructions\n" +
      "• Scene-by-scene assembly steps (" + scenes.length + " scenes)\n" +
      "• Exact text overlay styles with colors and fonts\n" +
      "• GRC closing argument card instructions\n" +
      "• CTA card instructions\n" +
      "• Audio settings\n" +
      "• YouTube export settings\n\n" +
      docUrl,
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 9 — Assembly", idea.id, "Google Doc creation failed", err.message);
    ui.alert("❌ Stage 9 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ── Helper: extract part from checkpoint description ──────────────────────────
function extractCheckpointPart(description, part) {
  const match = description.match(new RegExp(part + ":\\s*([^|]+)"));
  return match ? match[1].trim() : description;
}

// ── Helper: calculate scene duration from timestamps ─────────────────────────
function calculateDuration(currentTimestamp, nextTimestamp, format) {
  if (!nextTimestamp) {
    return format === "Short (< 90s)" ? "10" : "15";
  }
  try {
    const toSeconds = (ts) => {
      const parts = ts.toString().replace("–", "-").split(/[-:]/);
      if (parts.length >= 2) {
        return parseInt(parts[0]) * 60 + parseInt(parts[1]);
      }
      return 0;
    };
    const current = toSeconds(currentTimestamp.toString().split("–")[0].split("-")[0]);
    const next    = toSeconds(nextTimestamp.toString().split("–")[0].split("-")[0]);
    const diff    = next - current;
    return diff > 0 ? String(diff) : "10";
  } catch (e) {
    return "10";
  }
}

// ── Helper: get transition type based on scene type ───────────────────────────
function getTransition(sceneType, template) {
  const type = sceneType.toString().trim();
  if (type === "Text" || type === "Animation") return template.transitions.toText;
  if (type === "Minimal") return template.transitions.impact;
  if (type === "Checkpoint") return template.transitions.impact;
  if (type === "Timeline") return template.transitions.impact;
  return template.transitions.broll;
}

// ════════════════════════════════════════════════════════════════════════════════
// STAGE 8C — Export Google Veo Prompts
// Reads all scenes from Visual Library → converts to Veo-optimized prompts
// Writes prompt to VEO_PROMPT column + exports Google Doc for AI Studio
// ════════════════════════════════════════════════════════════════════════════════
function exportVeoPrompts() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Pull all scenes ───────────────────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    scenes.push({
      row            : i + 1,
      num            : visualData[i][COL_VISUAL.SCENE_NUM      - 1],
      type           : visualData[i][COL_VISUAL.SCENE_TYPE     - 1],
      description    : visualData[i][COL_VISUAL.DESCRIPTION    - 1],
      source         : visualData[i][COL_VISUAL.SOURCE         - 1],
      timestamp      : visualData[i][COL_VISUAL.TIMESTAMP      - 1],
      existingPrompt : visualData[i][COL_VISUAL.VEO_PROMPT     - 1]
    });
  }

  if (scenes.length === 0) {
    ui.alert("No scenes found for: " + idea.id + "\nRun Stage 4 first.");
    return;
  }

  // Check if VEO_PROMPT column already has values
  const hasPrompts = scenes.some(s => s.existingPrompt && s.existingPrompt.toString().trim() !== "");

  if (hasPrompts) {
    const response = ui.alert(
      "Veo Prompts Already Exist",
      "Some scenes already have Veo prompts.\n\nDo you want to regenerate all prompts using Claude?",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      // Export existing prompts to doc without regenerating
      exportVeoDoc(idea, master, scenes, visualSheet, false);
      // Still offer Veo API submission
      const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
      if (geminiKey) {
        const veoScenes = scenes.filter(s => {
          const type   = s.type.toString().trim();
          const source = s.source ? s.source.toString().trim() : "";
          return source === "Veo" ||
                 (source !== "CapCut" && source !== "Canva" &&
                  type !== "Text" && type !== "Checkpoint" &&
                  type !== "Timeline" && type !== "Animation" && type !== "Minimal");
        });
        if (veoScenes.length > 0) {
          const confirmVeo = ui.alert(
            "Submit to Veo API?",
            "Submit " + veoScenes.length + " scenes to Veo for generation?",
            ui.ButtonSet.YES_NO
          );
          if (confirmVeo === ui.Button.YES) {
            submitToVeoApi(idea, master, veoScenes, visualSheet, geminiKey);
          }
        }
      }
      return;
    }
  }

  ui.alert(
    "Stage 8C — Generating Veo Prompts",
    "Claude is converting " + scenes.length + " scenes into Google Veo prompts for: " + idea.company +
    "\nThis may take 20–30 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildVeoPromptRequest(idea, master, scenes);
    const raw    = callClaude(prompt, "stage_8_scenes");

    // Parse and write prompts back to Visual Library
    const veoBlocks = raw.match(/VEO_SCENE_\d+_START([\s\S]*?)VEO_SCENE_\d+_END/g) || [];

    veoBlocks.forEach(block => {
      const numMatch    = block.match(/VEO_SCENE_(\d+)_START/);
      const promptMatch = block.match(/PROMPT:\s*([\s\S]*?)(?=SETTINGS:|VEO_SCENE_\d+_END)/);
      const sceneNum    = numMatch    ? parseInt(numMatch[1])    : 0;
      const veoPrompt   = promptMatch ? promptMatch[1].trim()   : "";

      if (sceneNum > 0 && veoPrompt) {
        const scene = scenes.find(s => parseInt(s.num) === sceneNum);
        if (scene) {
          visualSheet.getRange(scene.row, COL_VISUAL.VEO_PROMPT).clearDataValidations();
          visualSheet.getRange(scene.row, COL_VISUAL.VEO_PROMPT).setValue(veoPrompt);
          scene.existingPrompt = veoPrompt; // update for doc export
        }
      }
    });

    // Export to Google Doc first (always kept)
    exportVeoDoc(idea, master, scenes, visualSheet, true);

    // ── Now submit to Veo API if key is available ─────────────────────────
    const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");

    if (!geminiKey) {
      ui.alert(
        "✅ Prompts Generated — Veo API Not Connected",
        "Google Doc exported successfully.\n\n" +
        "To auto-generate video clips, add GEMINI_API_KEY to Script Properties.\n" +
        "Get your key at: aistudio.google.com/apikey",
        ui.ButtonSet.OK
      );
      return;
    }

    // Filter only scenes that should be generated by Veo
    // Skip CapCut-built scenes (Text, Checkpoint, Timeline, Animation)
    const veoScenes = scenes.filter(s => {
      const type   = s.type.toString().trim();
      const source = s.source ? s.source.toString().trim() : "";
      return source === "Veo" ||
             (source !== "CapCut" && source !== "Canva" &&
              type !== "Text" && type !== "Checkpoint" &&
              type !== "Timeline" && type !== "Animation" &&
              type !== "Minimal");
    });

    if (veoScenes.length === 0) {
      ui.alert(
        "✅ Prompts Generated — No Veo Scenes",
        "Google Doc exported successfully.\n\n" +
        "No scenes require Veo generation for this video.\n" +
        "(All scenes are CapCut-built or Pexels stock.)",
        ui.ButtonSet.OK
      );
      return;
    }

    const confirmVeo = ui.alert(
      "Submit to Veo API?",
      "Found " + veoScenes.length + " scenes to generate with Veo.\n\n" +
      "Each clip takes 30–90 seconds.\n" +
      "Estimated total: ~" + Math.ceil(veoScenes.length * 1) + " minutes.\n\n" +
      "Clips will be saved to your Drive Scenes folder.\n" +
      "Proceed?",
      ui.ButtonSet.YES_NO
    );

    if (confirmVeo !== ui.Button.YES) return;

    submitToVeoApi(idea, master, veoScenes, visualSheet, geminiKey);

  } catch (err) {
    logError("Stage 8C — Veo Prompts", idea.id, "Claude API Error", err.message);
    ui.alert("❌ Stage 8C Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ── Build Claude prompt to convert all scenes to Veo prompts ─────────────────
function buildVeoPromptRequest(idea, master, scenes) {
  const sceneList = scenes.map(s =>
    `Scene ${s.num} [${s.timestamp}] — ${s.type}\nDescription: ${s.description}`
  ).join("\n\n");

  return `
You are converting GovernX video scene descriptions into optimized Google Veo 3 prompts 
for use in Google AI Studio.

CONTENT:
Company/Topic : ${idea.company}
Title         : ${master ? master.title : idea.company}
Discipline    : ${master ? master.discipline : "GRC"}
Visual Tone   : Dark, cinematic, documentary style. Professional. No text, no logos.

VEO PROMPT RULES:
- Written as a professional cinematographer's instruction
- Include: subject, camera movement, lighting, mood, color palette
- NEVER include: text, words, letters, numbers, logos, watermarks, brand names
- For Text/Animation scenes: describe an abstract or atmospheric visual instead
- For Stock scenes: describe the exact cinematic shot to recreate or find
- For AI Generated scenes: describe the atmospheric/symbolic visual
- Maximum 2 sentences per prompt
- End every prompt with: "No text, no watermarks, no logos. Cinematic 4K."

SCENES TO CONVERT:
${sceneList}

Return each scene in EXACTLY this format:

VEO_SCENE_1_START
PROMPT: [optimized Veo prompt — 1-2 sentences + "No text, no watermarks, no logos. Cinematic 4K."]
SETTINGS:
  Model: Veo 3.1
  Duration: 8s
  Frame rate: 24 fps
  Resolution: 4K
  Aspect ratio: 16:9
VEO_SCENE_1_END

[repeat for each scene]
`;
}

// ── Export Veo prompts to Google Doc ─────────────────────────────────────────
function exportVeoDoc(idea, master, scenes, visualSheet, regenerated) {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const divider = "═".repeat(60);
  const line    = "─".repeat(40);

  let doc = "";

  // Header
  doc += divider + "\n";
  doc += "GOVERNX — GOOGLE VEO PROMPT PACK\n";
  doc += divider + "\n";
  doc += "Content ID  : " + idea.id + "\n";
  doc += "Title       : " + (master ? master.title : idea.company) + "\n";
  doc += "Company     : " + idea.company + "\n";
  doc += "Discipline  : " + (master ? master.discipline : "") + "\n";
  doc += "Total scenes: " + scenes.length + "\n";
  doc += "Generated   : " + new Date().toLocaleString() + "\n";
  doc += divider + "\n\n";

  // Instructions
  doc += "HOW TO USE IN GOOGLE AI STUDIO\n";
  doc += line + "\n";
  doc += "1. Go to aistudio.google.com\n";
  doc += "2. Select 'Video' → 'Veo 3.1 Generate'\n";
  doc += "3. Link your API key if prompted\n";
  doc += "4. For each scene below:\n";
  doc += "   a. Copy the PROMPT text\n";
  doc += "   b. Paste into AI Studio prompt field\n";
  doc += "   c. Set the SETTINGS as shown\n";
  doc += "   d. Click Run\n";
  doc += "   e. Download the clip\n";
  doc += "   f. Save to your Drive Scenes folder\n";
  doc += "   g. Name it: " + idea.id + " — Scene-[XX] — Veo.mp4\n\n";
  doc += "RECOMMENDED ORDER: Start with the most important scenes first.\n";
  doc += "Compare results with your KlingAI clips to pick the best.\n\n";

  // Scene prompts
  doc += divider + "\n";
  doc += "SCENE PROMPTS\n";
  doc += divider + "\n\n";

  scenes.forEach(scene => {
    const sceneNum  = String(scene.num).padStart(2, "0");
    const veoPrompt = scene.existingPrompt
      ? scene.existingPrompt.toString().trim()
      : "(No prompt generated — re-run Stage 8C)";

    doc += "────────────────────────────────────────\n";
    doc += "SCENE " + sceneNum + " | " + scene.timestamp + " | " + scene.type.toUpperCase() + "\n";
    doc += "────────────────────────────────────────\n\n";
    doc += "ORIGINAL DESCRIPTION:\n";
    doc += scene.description + "\n\n";
    doc += "VEO PROMPT (copy this):\n";
    doc += "┌─────────────────────────────────────\n";
    doc += "│ " + veoPrompt.split("\n").join("\n│ ") + "\n";
    doc += "└─────────────────────────────────────\n\n";
    doc += "AI STUDIO SETTINGS:\n";
    doc += "  Model      : Veo 3.1\n";
    doc += "  Duration   : 8s\n";
    doc += "  Frame rate : 24 fps\n";
    doc += "  Resolution : 4K\n";
    doc += "  Aspect     : 16:9\n\n";
    doc += "After generating — save clip as:\n";
    doc += "  " + idea.id + " — Scene-" + sceneNum + " — Veo.mp4\n\n";
  });

  // Comparison notes
  doc += divider + "\n";
  doc += "QUALITY COMPARISON NOTES\n";
  doc += line + "\n\n";
  doc += "After generating all Veo clips, compare with KlingAI clips:\n\n";
  doc += "PICK VEO IF:\n";
  doc += "• Scene requires cinematic realism (people, environments, natural movement)\n";
  doc += "• KlingAI clip had text artifacts or garbled visuals\n";
  doc += "• You need longer duration (Veo supports 8s vs KlingAI 5s)\n";
  doc += "• 4K quality is important for this scene\n\n";
  doc += "PICK KLING IF:\n";
  doc += "• Scene is atmospheric/abstract (KlingAI handles mood well)\n";
  doc += "• Veo clip doesn't match the documentary feel\n";
  doc += "• KlingAI clip is already Ready in your Visual Library\n\n";
  doc += "PICK PEXELS IF:\n";
  doc += "• Real-world footage exists for this scene\n";
  doc += "• Neither AI tool produced satisfactory results\n\n";

  doc += divider + "\n";
  doc += "END OF VEO PROMPT PACK — " + idea.id + "\n";
  doc += divider + "\n";

  // Save to content folder
  try {
    const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
    const docTitle      = idea.id + " — " + idea.company + " — Veo Prompts";

    // Delete existing if any
    const existing = contentFolder.getFilesByName(docTitle);
    while (existing.hasNext()) existing.next().setTrashed(true);

    const newDoc = DocumentApp.create(docTitle);
    newDoc.getBody().setAttributes({ [DocumentApp.Attribute.FONT_FAMILY]: "Courier New" });
    newDoc.getBody().setText(doc);
    newDoc.saveAndClose();

    const file = DriveApp.getFileById(newDoc.getId());
    contentFolder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);

    const docUrl = newDoc.getUrl();

    ui.alert(
      "✅ Stage 8C Complete — Veo Prompts Exported",
      "Saved to Google Drive:\n" +
      "📁 " + DRIVE_FOLDER_NAME + " / " + idea.id + " — " + idea.company + "\n" +
      "📄 " + docTitle + "\n\n" +
      scenes.length + " scene prompts ready for Google AI Studio.\n\n" +
      "Each prompt includes:\n" +
      "• Optimized Veo 3.1 prompt\n" +
      "• Exact AI Studio settings (8s, 24fps, 4K, 16:9)\n" +
      "• File naming convention\n" +
      "• Quality comparison guide\n\n" +
      "VEO_PROMPT column in Visual Library also updated.\n\n" +
      docUrl,
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 8C — Export", idea.id, "Google Doc creation failed", err.message);
    ui.alert("❌ Stage 8C Export Failed", err.message, ui.ButtonSet.OK);
  }
}

// ════════════════════════════════════════════════════════════════════════════════
// VEO API HELPERS — Submit scenes to Google Veo via Gemini API
// ════════════════════════════════════════════════════════════════════════════════

function submitToVeoApi(idea, master, veoScenes, visualSheet, geminiKey) {
  const ui           = SpreadsheetApp.getUi();
  const scenesFolder = getOrCreateScenesFolder(idea.id, idea.company);
  let   success = 0, failed = 0;

  veoScenes.forEach(scene => {
    try {
      const veoPrompt = scene.existingPrompt
        ? scene.existingPrompt.toString().trim()
        : buildVeoScenePrompt(scene, idea);

      Logger.log("Submitting Scene " + scene.num + " to Veo: " + veoPrompt.substring(0, 80));

      // ── Step 1: Submit generation request ──────────────────────────────────
      const submitUrl = GEMINI_API_BASE + "/models/" + VEO_MODEL +
                        ":predictLongRunning?key=" + geminiKey;

      const submitPayload = {
        instances: [{
          prompt: veoPrompt
        }],
        parameters: {
          aspectRatio    : "16:9",
          durationSeconds: 8,
          fps            : 24,
          resolution     : "1080p",
          negativePrompt : "text, watermark, logo, captions, subtitles, words, letters, " +
                           "numbers, signs, blurry, low quality, amateur"
        }
      };

      const submitResponse = UrlFetchApp.fetch(submitUrl, {
        method            : "post",
        contentType       : "application/json",
        payload           : JSON.stringify(submitPayload),
        muteHttpExceptions: true
      });

      const submitCode = submitResponse.getResponseCode();
      const submitBody = JSON.parse(submitResponse.getContentText());

      if (submitCode === 403) {
        throw new Error(
          "PERMISSION_DENIED — Your API key does not have Veo access.\n" +
          "Try: (1) Use Veo 2 model in config.gs, or\n" +
          "(2) Request Veo access at ai.google.dev"
        );
      }

      if (submitCode !== 200) {
        throw new Error("Veo submit failed (" + submitCode + "): " + JSON.stringify(submitBody));
      }

      // Operation name for polling
      const operationName = submitBody.name;
      if (!operationName) {
        throw new Error("No operation name returned from Veo API");
      }

      Logger.log("Scene " + scene.num + " submitted. Operation: " + operationName);

      // ── Step 2: Poll for completion ─────────────────────────────────────────
      let videoUrl = null;

      for (let poll = 1; poll <= VEO_MAX_POLLS; poll++) {
        Utilities.sleep(VEO_POLL_INTERVAL);

        const pollUrl = GEMINI_API_BASE + "/" + operationName + "?key=" + geminiKey;
        const pollResponse = UrlFetchApp.fetch(pollUrl, {
          method            : "get",
          muteHttpExceptions: true
        });

        const pollBody = JSON.parse(pollResponse.getContentText());
        Logger.log("Scene " + scene.num + " poll " + poll + ": done=" + pollBody.done);

        if (pollBody.done === true) {
          // Extract video URL — Veo API returns several possible structures:
          // Structure A: pollBody.response.predictions[0].bytesBase64Encoded
          // Structure B: pollBody.response.predictions[0].gcsUri / .videoUri
          // Structure C: pollBody.response.predictions[0].video.uri  (Veo 2 Vertex AI)
          // Structure D: pollBody.response.videos[0].uri             (some Veo 2 variants)
          const predictions = pollBody.response && pollBody.response.predictions;
          const videosArr   = pollBody.response && pollBody.response.videos;

          if (predictions && predictions.length > 0) {
            const pred = predictions[0];
            if (pred.bytesBase64Encoded) {
              const videoBytes = Utilities.base64Decode(pred.bytesBase64Encoded);
              const videoBlob  = Utilities.newBlob(videoBytes, "video/mp4");
              const fileName   = idea.id + " — Scene-" +
                                 String(scene.num).padStart(2, "0") + " — Veo.mp4";
              const existing = scenesFolder.getFilesByName(fileName);
              while (existing.hasNext()) existing.next().setTrashed(true);
              const savedFile = scenesFolder.createFile(videoBlob.setName(fileName));
              videoUrl = "https://drive.google.com/file/d/" + savedFile.getId() + "/view";
              Logger.log("Scene " + scene.num + " Veo clip saved (base64): " + videoUrl);
            } else if (pred.video && pred.video.uri) {
              videoUrl = pred.video.uri;   // Vertex AI Veo 2 structure
            } else if (pred.gcsUri || pred.videoUri) {
              videoUrl = pred.gcsUri || pred.videoUri;
            }
          }

          // Fallback: videos array format
          if (!videoUrl && videosArr && videosArr.length > 0) {
            videoUrl = videosArr[0].uri || videosArr[0].url || videosArr[0].gcsUri || null;
          }

          if (!videoUrl) {
            // Log the full response for debugging before throwing
            Logger.log("Veo render completed but video URL not found in response. " +
              "Full structure: " + JSON.stringify(pollBody).substring(0, 500));
            throw new Error(
              "Veo render completed but video URL not found in response.\n" +
              "Check execution log for full response structure.\n" +
              "Operation: " + operationName
            );
          }

          break;
        }

        if (pollBody.error) {
          throw new Error("Veo processing error: " + JSON.stringify(pollBody.error));
        }
      }

      if (!videoUrl) {
        throw new Error(
          "Timed out waiting for Veo Scene " + scene.num +
          " after " + (VEO_MAX_POLLS * VEO_POLL_INTERVAL / 60000).toFixed(1) + " minutes. " +
          "Operation: " + operationName
        );
      }

      // ── Step 3: Write URL to Visual Library ─────────────────────────────────
      visualSheet.getRange(scene.row, COL_VISUAL.AI_CLIP_URL).setValue(videoUrl);
      visualSheet.getRange(scene.row, COL_VISUAL.STATUS     ).setValue("Ready");
      success++;

    } catch (err) {
      failed++;
      logError("Stage 8C — Veo Scene " + scene.num, idea.id, "Veo API Error", err.message);
      visualSheet.getRange(scene.row, COL_VISUAL.STATUS).setValue("Error");
      Logger.log("Scene " + scene.num + " Veo failed: " + err.message);
    }
  });

  // ── Summary alert ────────────────────────────────────────────────────────────
  const scenesFolderUrl = "https://drive.google.com/drive/folders/" + scenesFolder.getId();

  ui.alert(
    success > 0 && failed === 0 ? "✅ Veo Generation Complete" :
    failed > 0 ? "⚠️ Veo Generation — Partial Results" : "❌ Veo Generation Failed",
    "Results for: " + idea.id + "\n\n" +
    "✅ Generated: " + success + " clips\n" +
    (failed > 0 ? "❌ Failed: " + failed + " clips (see Error Log)\n\n" : "\n") +
    "Clip links saved in Visual Library → AI Clip URL column.\n\n" +
    "📁 Scenes folder:\n" + scenesFolderUrl,
    ui.ButtonSet.OK
  );
}

// ── Build Veo prompt for a single scene (fallback if no existing prompt) ──────
function buildVeoScenePrompt(scene, idea) {
  return "Cinematic documentary B-roll footage for " + idea.company + " story. " +
    scene.description + " " +
    "Style: dark professional tone, cinematic 4K, sharp focus, no camera shake. " +
    "CRITICAL: NO text, NO words, NO letters, NO numbers, NO signs, NO logos, " +
    "NO watermarks, NO captions anywhere in frame. Pure visual footage only.";
}


// ════════════════════════════════════════════════════════════════════════════════
// STAGE 10 — Generate YouTube Metadata (canonical version)
// Uses the full-featured buildYouTubeMetadataPrompt (5-param) and
// writeYouTubeMetadata (4-param) defined above — writes all 11 metadata fields.
// ════════════════════════════════════════════════════════════════════════════════
function generateYouTubeMetadata() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Pull script data ──────────────────────────────────────────────────────
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let   script      = null;

  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      script = {
        hook      : scriptData[i][COL_SCRIPT.HOOK             - 1],
        voiceover : scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1],
        closing   : scriptData[i][COL_SCRIPT.GRC_BPR_CLOSING  - 1],
        cta       : scriptData[i][COL_SCRIPT.CALL_TO_ACTION   - 1],
        format    : scriptData[i][COL_SCRIPT.TARGET_FORMAT    - 1],
        sections  : scriptData[i][COL_SCRIPT.SECTIONS         - 1]
      };
      break;
    }
  }

  if (!script) {
    ui.alert("No script found for: " + idea.id + "\nRun Stage 3 first.");
    return;
  }

  // ── Pull scene timestamps for chapter markers ─────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const scenes      = [];
  if (visualSheet) {
    const visualData = visualSheet.getDataRange().getValues();
    for (let i = 1; i < visualData.length; i++) {
      if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
      const sceneType = visualData[i][COL_VISUAL.SCENE_TYPE - 1].toString().trim();
      const timestamp = visualData[i][COL_VISUAL.TIMESTAMP  - 1].toString().trim();
      const desc      = visualData[i][COL_VISUAL.DESCRIPTION- 1].toString().trim();
      if (["Text", "Checkpoint", "Infographic", "Timeline"].includes(sceneType) && timestamp && desc) {
        scenes.push({
          num       : visualData[i][COL_VISUAL.SCENE_NUM - 1],
          type      : sceneType,
          desc      : desc.substring(0, 60),
          timestamp : timestamp
        });
      }
    }
  }

  // ── Pull research sources (null-safe) ─────────────────────────────────────
  // getResearchSources() returns null when all sources are rejected.
  // Stage 10 is informational — use empty array rather than blocking.
  const sourcesRaw = getResearchSources(idea.id);
  const sources    = sourcesRaw || [];

  ui.alert(
    "Stage 10 — Generating YouTube Metadata",
    "Claude is generating all YouTube metadata for: " + idea.company +
    "\n\nThis includes 3 title variants, description, tags, chapters,\n" +
    "hashtags, pinned comment, end screen CTA, and thumbnail brief.\n\n" +
    "This may take 20–30 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const prompt = buildYouTubeMetadataPrompt(idea, master, script, scenes, sources);
    const raw    = callClaude(prompt, "stage_10_youtube");

    writeYouTubeMetadata(idea.id, raw, master, script, scenes);
    updatePipelineStatus_(idea.id, "S10", "✅");

    ui.alert(
      "✅ Stage 10 Complete — YouTube Metadata Generated",
      "YouTube Metadata tab has been populated for: " + idea.id + "\n\n" +
      "Contains:\n" +
      "• 3 title variants (A/B test ready)\n" +
      "• Full SEO-optimized description with all sources\n" +
      "• Tags (40+ keywords)\n" +
      "• Chapter timestamps (from the assembled film)\n" +
      "• Hashtags\n" +
      "• Pinned first comment\n" +
      "• End screen CTA\n" +
      "• Thumbnail brief\n\n" +
      "REVIEW the YouTube Metadata tab — especially the title and description —\n" +
      "then run Stage 11 to upload. (Upload is no longer automatic: a publish\n" +
      "step should never fire off metadata generation without your review.)",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S10", "❌");
    logError("Stage 10 — YouTube Metadata", idea.id, "API Error", err.message);
    ui.alert("❌ Stage 10 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}

// ── Build YouTube metadata prompt ─────────────────────────────────────────────
function buildYouTubeMetadataPrompt(idea, master, script, scenes, sources) {

  const sceneList = scenes.map(s =>
    "[" + s.timestamp + "] Scene " + s.num + " (" + s.type + "): " + s.desc
  ).join("\n");

  const sourceList = (sources || []).slice(0, 5).map((s, i) =>
    (i + 1) + ". " + s.details
  ).join("\n");

  return `
You are generating complete YouTube metadata for a GovernX video.

CHANNEL: GovernX — reverse-engineers leadership decisions to reveal GRC and BPR urgency.
AUDIENCE: C-Suite executives, board members, compliance professionals, business students.
TONE: Sharp, documentary, authoritative. Not clickbait — intelligent hooks.

CONTENT BRIEF:
Company/Topic  : ${idea.company}
Title (draft)  : ${master ? master.title : idea.company}
Hook           : ${script.hook}
Core Insight   : ${master ? master.coreInsight : ""}
Discipline     : ${master ? master.discipline : ""}
Series         : ${idea.series}
Format         : ${script.format}
GRC Closing    : ${script.closing}
CTA            : ${script.cta}
Language       : ${idea.language}

SCENE TIMESTAMPS:
${sceneList || "No scene timestamps available"}

KEY SOURCES USED:
${sourceList || "No sources available"}

GENERATE the following in EXACTLY this format:

TITLE_A: [Primary title — sharp, specific, includes company name and the failure — under 60 chars]
TITLE_B: [A/B variant — leads with the number or stat — under 60 chars]
TITLE_C: [A/B variant — question format "Why did..." or "How did..." — under 60 chars]

DESCRIPTION_START
[Full YouTube description — 150-200 words]
[Line 1-2: Hook — the visible outcome, specific stat]
[Line 3-5: What the video reveals — the reverse-engineering angle]
[Line 6-8: The GRC/BPR lesson and why it matters]
[Line 9-10: Series context and subscribe CTA]

📚 SOURCES:
[List top 3 sources with titles]

🔗 GovernX Series: ${idea.series}
[End the description with 4-5 hashtags. The first three render above the title, so lead with the SPECIFIC ones — company, person, core topic (e.g. #Nissan #CarlosGhosn #CorporateGovernance) — then #GovernX #GRC. Keep them consistent with the HASHTAGS field below.]
DESCRIPTION_END

TAGS: [25-30 comma-separated YouTube tags, MOST SPECIFIC FIRST: company name, the people involved, and the exact event; THEN GRC/governance terms; THEN broader industry terms. Keep the ENTIRE list under 480 characters (YouTube drops anything past 500). No vague filler — avoid "risk management", "corporate ethics", "leadership", "organizational risk".]

CHAPTERS_START
[Chapter timestamps — format: MM:SS Chapter Title]
[Always start with 0:00 Introduction]
[End with final timestamp GRC Lesson]
CHAPTERS_END

HASHTAGS: [8-10 hashtags. YouTube shows ONLY THE FIRST THREE above the title, so make those three the strongest and most-searched — the company, the person, and the core topic (e.g. #Nissan #CarlosGhosn #CorporateGovernance). Put #GovernX and any generic tags AFTER the first three. Do NOT force #BPR into the top three.]

FIRST_COMMENT: [Pinned comment — 2-3 sentences. Pose a question to drive engagement. Reference the GovernX reverse-engineering method. Ask audience what company they want analyzed next.]

END_SCREEN: [End-screen CTA — ONE concrete, punchy line that echoes the film's closing thesis, then a subscribe push. Mirror the GRC Closing above and the channel's signature framing "Every collapse has an architecture." Example pattern: "Every collapse has an architecture. Subscribe to GovernX — learn to see it before it breaks." Do NOT invent abstract phrasings like "every satisfactory outcome"; keep it about collapse/failure and tie it to this specific case.]

THUMBNAIL_BRIEF: [Specific visual brief: TEXT overlay (max 5 words, high contrast), background image concept, color treatment, emotional hook. Format: "TEXT: [words] | IMAGE: [concept] | COLOR: [treatment] | HOOK: [why it stops scrolling]"]
`;
}

/* ── Build a YouTube-valid chapter list from REAL scene start times ────────────
   `scenes` carry `timestamp` = "M:SS" written by Stage 9C from the assembled
   film. YouTube requires: first chapter at 0:00, ≥3 chapters, each ≥10s after the
   previous, ascending. One chapter per scene is too many, so scenes are merged
   until at least MIN_GAP seconds have passed. `claudeChapters` is used only to
   borrow a nicer title when one lines up. */
function buildYouTubeChapters_(scenes, claudeChapters) {
  const toSec = function (mmss) {
    const m = String(mmss || "").trim().match(/^(\d+):(\d{1,2})$/);
    return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : null;
  };
  const fmt = function (sec) { return Math.floor(sec / 60) + ":" + String(sec % 60).padStart(2, "0"); };

  // Only scenes with a real timestamp, sorted, deduped.
  const withTime = (scenes || [])
    .map(function (s) { return { sec: toSec(s.timestamp), type: s.type, desc: String(s.desc || "").trim() }; })
    .filter(function (s) { return s.sec !== null; })
    .sort(function (a, b) { return a.sec - b.sec; });

  const MIN_GAP = 18;   // seconds between chapters (comfortably over YouTube's 10s floor)
  const chapterTitle = function (s) {
    // A checkpoint's own event line is the best title; otherwise trim the desc.
    // Split only on SENTENCE punctuation (". ", ": ", " — ") so a decimal figure
    // like "9.078 billion yen" is never chopped to "9".
    let t = s.desc.replace(/^\[[^\]]*\]\s*/, "").replace(/^[A-Z_]+\s+—\s+/, "").trim();
    t = t.split(/\.\s|:\s|\s—\s/)[0].trim();
    return (t || s.type || "Chapter").slice(0, 48);
  };

  const out = [{ sec: 0, title: "Introduction" }];
  let last = 0;
  withTime.forEach(function (s) {
    if (s.sec - last < MIN_GAP) return;      // too close to the previous chapter
    if (s.sec < 10) return;                  // never a chapter before 0:10
    out.push({ sec: s.sec, title: chapterTitle(s) });
    last = s.sec;
  });

  // YouTube needs at least 3 chapters; if merging left too few, relax the gap once.
  if (out.length < 3 && withTime.length) {
    out.length = 1; last = 0;
    withTime.forEach(function (s) {
      if (s.sec - last < 11 || s.sec < 10) return;
      out.push({ sec: s.sec, title: chapterTitle(s) });
      last = s.sec;
    });
  }

  // Still under 3 — almost always because Stage 10 ran BEFORE Stage 9C wrote the
  // per-scene timestamps, so no real times exist. A lone "0:00 Introduction" does
  // nothing on YouTube (it needs ≥3 to render chapters at all) and looks broken.
  // Return blank instead; writeYouTubeMetadata flags the reason in the Note column.
  if (out.length < 3) return "";

  return out.map(function (c) { return fmt(c.sec) + " " + c.title; }).join("\n");
}

/* Cap a comma-separated tag list to whole tags that fit YouTube's 500-char limit.
   YouTube silently drops everything past 500, so the tail tags are lost anyway —
   trimming here keeps the highest-value tags (the model is told to put specific
   ones first). Headroom default 480. */
function capTags_(tagString, maxChars) {
  const cap = maxChars || 480;
  const tags = String(tagString || "").split(",").map(function (t) { return t.trim(); }).filter(Boolean);
  const kept = [];
  let len = 0;
  for (let i = 0; i < tags.length; i++) {
    const add = (kept.length ? 2 : 0) + tags[i].length;   // ", " joins
    if (len + add > cap) break;
    kept.push(tags[i]);
    len += add;
  }
  return kept.join(", ");
}

// ── Write YouTube metadata to dedicated tab ───────────────────────────────────
function writeYouTubeMetadata(contentId, raw, master, script, scenes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getLine = (field) => {
    const match = raw.match(new RegExp(field + ":\\s*([^\\n\\r]*)"));
    if (!match) return "";
    const val = match[1].trim();
    return val === "NONE" || val === "" ? "" : val;
  };

  const getBlock = (startTag, endTag) => {
    const match = raw.match(new RegExp(startTag + "([\\s\\S]*?)" + endTag));
    return match ? match[1].trim() : "";
  };

  // Chapters are built from the REAL per-scene start times that Stage 9C wrote to
  // the Visual Library TIMESTAMP column after assembly. Claude cannot know the
  // exact film timing, so its guessed timestamps are ignored — its chapter TITLES
  // are only used as a naming hint. YouTube's rules are enforced here:
  //   • first chapter MUST be 0:00
  //   • at least 3 chapters, each at least 10 seconds after the previous
  //   • ascending order
  // Too many chapters (one per scene) is both ugly and rejected, so scenes are
  // merged to roughly one chapter per major beat.
  const chapters = buildYouTubeChapters_(scenes, getBlock("CHAPTERS_START", "CHAPTERS_END") || getLine("CHAPTERS"));

  // If chapters came back blank, the scene timestamps weren't there yet — tell the
  // user exactly what to do rather than shipping a metadata row with no chapters.
  const chapterNote = chapters ? "" :
    "⚠ No chapters: per-scene timestamps missing. Run Stage 9C (Assemble Film) FIRST so timestamps are written, then re-run Stage 10.";

  // ── Get or create YouTube Metadata tab ───────────────────────────────────
  let sheet = ss.getSheetByName(SHEET.YOUTUBE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET.YOUTUBE);
    const headers = ["ID", "Title A", "Title B", "Title C", "Description",
                     "Tags", "Chapters", "Hashtags", "First Comment",
                     "End Screen", "Thumbnail Brief", "Status", "Note"];
    headers.forEach((h, i) => {
      sheet.getRange(1, i + 1)
        .setValue(h).setFontWeight("bold")
        .setBackground("#1E293B").setFontColor("#FFFFFF").setFontFamily("Montserrat");
    });
    [160,280,280,280,400,300,200,200,300,200,300,100,200].forEach((w, i) =>
      sheet.setColumnWidth(i + 1, w));
    sheet.setFrozenRows(1);
  }

  // ── Upsert row ────────────────────────────────────────────────────────────
  let targetRow = -1;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_YOUTUBE.ID - 1].toString().trim() === contentId) {
      targetRow = i + 1; break;
    }
  }
  if (targetRow === -1) targetRow = sheet.getLastRow() + 1;

  sheet.setRowHeight(targetRow, 120);
  for (let c = 1; c <= 13; c++) sheet.getRange(targetRow, c).clearDataValidations();

  sheet.getRange(targetRow, COL_YOUTUBE.ID             ).setValue(contentId);
  sheet.getRange(targetRow, COL_YOUTUBE.TITLE_A        ).setValue(getLine("TITLE_A")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.TITLE_B        ).setValue(getLine("TITLE_B")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.TITLE_C        ).setValue(getLine("TITLE_C")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.DESCRIPTION    ).setValue(getBlock("DESCRIPTION_START", "DESCRIPTION_END")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.TAGS           ).setValue(capTags_(getLine("TAGS"), 480)).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.CHAPTERS       ).setValue(chapters).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.HASHTAGS       ).setValue(getLine("HASHTAGS")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.FIRST_COMMENT  ).setValue(getLine("FIRST_COMMENT")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.END_SCREEN     ).setValue(getLine("END_SCREEN")).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.THUMBNAIL_BRIEF).setValue(
    getLine("THUMBNAIL_BRIEF") || (master ? master.thumbnailBrief || "" : "")
  ).setWrap(true);
  sheet.getRange(targetRow, COL_YOUTUBE.STATUS         ).setValue("Draft");
  sheet.getRange(targetRow, COL_YOUTUBE.NOTE           ).setValue(chapterNote).setWrap(true);

  const bg = (targetRow % 2 === 0) ? "#F8F9FA" : "#FFFFFF";
  sheet.getRange(targetRow, 1, 1, 13).setBackground(bg);

  Logger.log("YouTube Metadata written for: " + contentId);
  SpreadsheetApp.flush();
}
// ════════════════════════════════════════════════════════════════════════════════
// autoResolveErrorLog
// Marks Error Log entries as resolved when the associated scene reaches Done OR Ready.
// Call this after any batch status update (Stage 8D, 9B completions).
// ════════════════════════════════════════════════════════════════════════════════
function autoResolveErrorLog(contentId) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet  = ss.getSheetByName(SHEET.ERROR_LOG);
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  if (!errorSheet || !visualSheet) return;

  // Build set of scene numbers that are Done or Ready (lifecycle ends at Ready)
  const visualData = visualSheet.getDataRange().getValues();
  const resolvedScenes = new Set();
  for (let i = 1; i < visualData.length; i++) {
    const id     = visualData[i][COL_VISUAL.ID       - 1].toString().trim();
    const scNum  = visualData[i][COL_VISUAL.SCENE_NUM - 1].toString().trim();
    const status = visualData[i][COL_VISUAL.STATUS    - 1].toString().trim();
    if (id === contentId && (status === "Done" || status === "Ready")) {
      resolvedScenes.add(scNum);
    }
  }
  if (resolvedScenes.size === 0) return;

  const errorData = errorSheet.getDataRange().getValues();
  let resolved = 0;
  for (let i = 1; i < errorData.length; i++) {
    const errId      = errorData[i][COL_ERROR.ID       - 1].toString().trim();
    const errStage   = errorData[i][COL_ERROR.STAGE    - 1].toString();
    const isResolved = errorData[i][COL_ERROR.RESOLVED - 1];
    if (errId !== contentId) continue;
    if (isResolved === true || isResolved === "TRUE") continue;
    const scMatch = errStage.match(/Scene\s+([^\s]+)/i);
    if (scMatch && resolvedScenes.has(scMatch[1])) {
      errorSheet.getRange(i + 1, COL_ERROR.RESOLVED).setValue(true);
      errorSheet.getRange(i + 1, 1, 1, 6).setBackground("#d4edda");
      resolved++;
    }
  }
  if (resolved > 0) Logger.log("Auto-resolved " + resolved + " error entries for: " + contentId);
}