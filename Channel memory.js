/* ============================================================================
   Channel_Memory.gs — GovernX Content OS
   M8 — Channel Brain

   WHAT THIS IS:
   The persistent learning layer for GovernX. After every video completes
   Stage 10, the Channel Brain writes a structured memory record to a JSON
   file in Google Drive. Before Stage 1 runs on the next video, it reads
   the last 5 records and injects them into buildMasterContentPrompt() so
   Claude knows what worked, what the audience responded to, and what
   patterns to avoid or repeat.

   The system gets smarter with every video published.

   MEMORY FILE:
   Stored as "GovernX_Channel_Memory.json" inside the root Drive folder
   "GovernX Production Packages". One JSON array of memory records,
   appended after each video.

   TWO FUNCTIONS:
   writeChannelMemory(contentId) — called after Stage 10 completes
   readChannelMemory()           — called inside buildMasterContentPrompt()

   MENU INTEGRATION:
   Add to Menu.gs after Stage 10:
     .addItem("🧠 Write Channel Memory", "writeChannelMemoryForSelected")

   PIPELINE INTEGRATION:
   In Pipeline.gs → buildMasterContentPrompt(idea):
   Add at the top of the function:
     const memoryContext = readChannelMemory();
   Then inject ${memoryContext} into the prompt template.
   ============================================================================ */


// ── Memory file name ──────────────────────────────────────────────────────────
const MEMORY_FILE_NAME = "GovernX_Channel_Memory.json";
const MEMORY_MAX_RECORDS = 20;   // Keep last 20 videos in memory
const MEMORY_INJECT_COUNT = 5;   // Inject last 5 into Stage 1 prompt


// ══════════════════════════════════════════════════════════════════════════════
// READ CHANNEL MEMORY
// Called from buildMasterContentPrompt() in Pipeline.gs
// Returns a formatted string ready to inject into the Stage 1 prompt
// Returns empty string gracefully if no memory exists yet
// ══════════════════════════════════════════════════════════════════════════════
function readChannelMemory() {
  try {
    const file    = getMemoryFile(false); // false = don't create if missing
    if (!file) return "";                 // no memory yet — first video

    const content = file.getBlob().getDataAsString();
    if (!content || content.trim() === "") return "";

    const records = JSON.parse(content);
    if (!Array.isArray(records) || records.length === 0) return "";

    // Take the most recent N records
    const recent = records.slice(-MEMORY_INJECT_COUNT);

    // Build the context block to inject into the prompt
    let ctx = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CHANNEL MEMORY — LEARN FROM PREVIOUS GOVERNX VIDEOS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
The following records summarize the last ${recent.length} GovernX video(s).
Use this memory to:
- Avoid repeating the same domain, industry, or GRC angle back-to-back
- Build on hook styles and narrative patterns that scored highly
- Apply lessons from QA Critic scores to strengthen this new video
- Ensure geographic and domain diversity across the content calendar
- Reference what topics the audience has already seen

`;

    recent.forEach((rec, idx) => {
      ctx += `MEMORY RECORD ${idx + 1} — ${rec.contentId} | ${rec.company}\n`;
      ctx += `─────────────────────────────────────────\n`;
      ctx += `Domain         : ${rec.domain} / ${rec.industry}\n`;
      ctx += `Discipline     : ${rec.discipline}\n`;
      ctx += `GRC Angle      : ${rec.grcAngle}\n`;
      ctx += `Hook           : ${rec.hook}\n`;
      ctx += `Format         : ${rec.format}\n`;
      ctx += `Language       : ${rec.language}\n`;
      ctx += `Published      : ${rec.publishDate || "Not yet published"}\n`;

      // QA Critic scores
      if (rec.qaScores) {
        ctx += `QA Critic Scores:\n`;
        ctx += `  Hook Strength      : ${rec.qaScores.hookScore || "—"}/10\n`;
        ctx += `  Reverse Engineering: ${rec.qaScores.reverseScore || "—"}/10\n`;
        ctx += `  GRC Lesson Clarity : ${rec.qaScores.grcScore || "—"}/10\n`;
        ctx += `  Narrative Pacing   : ${rec.qaScores.pacingScore || "—"}/10\n`;
        ctx += `  Data Completeness  : ${rec.qaScores.dataCompletenessScore || "—"}/10\n`;
        ctx += `  Total              : ${rec.qaScores.totalScore || "—"}/70\n`;
        ctx += `  Verdict            : ${rec.qaScores.verdict || "—"}\n`;
      }

      // Director verdicts
      if (rec.directorVerdict) {
        ctx += `Director Verdict   : ${rec.directorVerdict}\n`;
        ctx += `Infographics Added : ${rec.infographicsAdded || 0}\n`;
      }

      // YouTube performance (if available)
      if (rec.performance && rec.performance.hasData) {
        ctx += `YouTube Performance:\n`;
        ctx += `  Views         : ${rec.performance.views || "—"}\n`;
        ctx += `  CTR           : ${rec.performance.ctr || "—"}%\n`;
        ctx += `  Retention     : ${rec.performance.retention || "—"}%\n`;
        ctx += `  Avg Watch Time: ${rec.performance.avgWatchTime || "—"}\n`;
        ctx += `  Perf Score    : ${rec.performance.perfScore || "—"}\n`;
      }

      // What worked / what to improve
      if (rec.lessonsLearned) {
        ctx += `Lessons Learned:\n`;
        ctx += `  What worked    : ${rec.lessonsLearned.worked || "—"}\n`;
        ctx += `  What to improve: ${rec.lessonsLearned.improve || "—"}\n`;
      }

      ctx += "\n";
    });

    ctx += `━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
END OF CHANNEL MEMORY — apply these lessons to the new video above.
`;

    return ctx;

  } catch (err) {
    Logger.log("Channel Memory read error (non-fatal): " + err.message);
    return ""; // memory failure is never blocking — pipeline continues
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// WRITE CHANNEL MEMORY
// Called after Stage 10 completes for a content ID
// Reads all available data from the pipeline tabs and writes a memory record
// ══════════════════════════════════════════════════════════════════════════════
function writeChannelMemoryForSelected() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    "🧠 Write Channel Memory",
    "Write memory record for: " + idea.id + " — " + idea.company + "\n\n" +
    "This captures:\n" +
    "• Content brief (domain, discipline, GRC angle)\n" +
    "• QA Critic scores (if Stage 3 was run with v2.0)\n" +
    "• Director verdict (if Stage 4.5 was run)\n" +
    "• YouTube performance (if data is in Publishing Tracker)\n\n" +
    "Run this after Stage 10 is complete.\nProceed?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  try {
    writeChannelMemory(idea.id, idea.company);
    ui.alert(
      "✅ Memory Written",
      "Channel Memory updated for: " + idea.id + "\n\n" +
      "This record will inform the next video's Stage 1 brief.",
      ui.ButtonSet.OK
    );
  } catch (err) {
    logError("Channel Memory", idea.id, "Write Error", err.message);
    ui.alert("❌ Memory Write Failed", err.message, ui.ButtonSet.OK);
  }
}


function writeChannelMemory(contentId, company) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const master = getMasterContent(contentId);

  // ── Build memory record ──────────────────────────────────────────────────
  const record = {
    contentId        : contentId,
    company          : company || contentId,
    writtenAt        : new Date().toISOString(),
    publishDate      : "",
    domain           : "",
    industry         : "",
    discipline       : "",
    grcAngle         : "",
    hook             : "",
    format           : "",
    language         : "",
    qaScores         : null,
    directorVerdict  : null,
    infographicsAdded: 0,
    performance      : { hasData: false },
    lessonsLearned   : null
  };

  // ── Pull from Master Content ────────────────────────────────────────────
  if (master) {
    record.domain     = master.domain     || "";
    record.industry   = master.industry   || "";
    record.discipline = master.discipline || "";
    record.grcAngle   = master.primaryAngle || master.reverseAngle || "";
    record.hook       = master.hook       || "";
    record.language   = master.language   || "";
  }

  // ── Pull format from Script Bank ────────────────────────────────────────
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  if (scriptSheet) {
    const scriptData = scriptSheet.getDataRange().getValues();
    for (let i = 1; i < scriptData.length; i++) {
      if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === contentId) {
        record.format = scriptData[i][COL_SCRIPT.TARGET_FORMAT - 1] || "";
        break;
      }
    }
  }

  // ── Pull QA Critic scores from Error Log (stored when evaluator ran) ────
  // QA scores are logged via logError when verdict is not APPROVED
  // For APPROVED scripts, we read the score from a dedicated memory store
  // Check for QA score in Script Bank NOTE column or Error Log
  record.qaScores = readQAScoresForContent(ss, contentId);

  // ── Pull Director verdict from Publishing Tracker Notes ─────────────────
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (pubSheet) {
    const pubData = pubSheet.getDataRange().getValues();
    for (let i = 1; i < pubData.length; i++) {
      if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() !== contentId) continue;

      record.publishDate = pubData[i][COL_PUBLISHING.PUBLISH_DATE - 1]
        ? new Date(pubData[i][COL_PUBLISHING.PUBLISH_DATE - 1]).toISOString().split("T")[0]
        : "";

      // YouTube performance data
      const views       = pubData[i][COL_PUBLISHING.VIEWS          - 1];
      const ctr         = pubData[i][COL_PUBLISHING.CTR            - 1];
      const retention   = pubData[i][COL_PUBLISHING.RETENTION      - 1];
      const avgWatch    = pubData[i][COL_PUBLISHING.AVG_WATCH_TIME - 1];
      const perfScore   = pubData[i][COL_PUBLISHING.PERF_SCORE     - 1];

      if (views || ctr || retention) {
        record.performance = {
          hasData     : true,
          views       : views       || 0,
          ctr         : ctr         || 0,
          retention   : retention   || 0,
          avgWatchTime: avgWatch    || "—",
          perfScore   : perfScore   || "—"
        };
      }

      // Extract Director verdict from Notes column
      const notes = pubData[i][COL_PUBLISHING.NOTES - 1].toString();
      const directorMatch = notes.match(/Director Report:/);
      if (directorMatch) {
        // Director pass was run — extract verdict from Error Log or mark as run
        record.directorVerdict = "Completed — see Director Report in Drive";
      }

      break;
    }
  }

  // ── Generate lessons learned using Claude ───────────────────────────────
  record.lessonsLearned = generateLessonsLearned(record);

  // ── Read existing memory file ────────────────────────────────────────────
  const memFile = getMemoryFile(true); // true = create if missing
  let   records = [];

  try {
    const existing = memFile.getBlob().getDataAsString();
    if (existing && existing.trim() !== "") {
      records = JSON.parse(existing);
      if (!Array.isArray(records)) records = [];
    }
  } catch (e) {
    records = []; // corrupt file — start fresh
  }

  // Remove any existing record for this contentId (upsert)
  records = records.filter(r => r.contentId !== contentId);

  // Append new record
  records.push(record);

  // Trim to max records (keep most recent)
  if (records.length > MEMORY_MAX_RECORDS) {
    records = records.slice(-MEMORY_MAX_RECORDS);
  }

  // ── Write back to Drive ──────────────────────────────────────────────────
  const json = JSON.stringify(records, null, 2);
  memFile.setContent(json);

  Logger.log("Channel Memory written for: " + contentId +
    " | Total records: " + records.length);
}


// ══════════════════════════════════════════════════════════════════════════════
// GENERATE LESSONS LEARNED
// Uses a lightweight Claude call to synthesize what the data means
// Only runs if there is performance data — otherwise returns a stub
// ══════════════════════════════════════════════════════════════════════════════
function generateLessonsLearned(record) {

  // If no meaningful data yet, return a stub
  const hasQA     = record.qaScores && record.qaScores.totalScore;
  const hasPerf   = record.performance && record.performance.hasData;
  const hasDir    = record.directorVerdict;

  if (!hasQA && !hasPerf) {
    return {
      worked  : "No performance data yet — check back after video publishes",
      improve : "Run Stage 3 with v2.0 to capture QA Critic scores"
    };
  }

  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY");
    if (!apiKey) throw new Error("No API key");

    const prompt = `
You are analyzing GovernX video performance data to extract concise lessons.
GovernX reverse-engineers leadership decisions to teach GRC and BPR.

VIDEO: ${record.contentId} — ${record.company}
Domain: ${record.domain} | Discipline: ${record.discipline}
Hook: ${record.hook}

${hasQA ? `QA CRITIC SCORES:
Hook Strength: ${record.qaScores.hookScore}/10
Reverse Engineering: ${record.qaScores.reverseScore}/10
GRC Lesson Clarity: ${record.qaScores.grcScore}/10
Narrative Pacing: ${record.qaScores.pacingScore}/10
Total: ${record.qaScores.totalScore}/70
Verdict: ${record.qaScores.verdict}` : "QA scores: Not available"}

${hasPerf ? `YOUTUBE PERFORMANCE:
Views: ${record.performance.views}
CTR: ${record.performance.ctr}%
Retention: ${record.performance.retention}%
Performance Score: ${record.performance.perfScore}` : "YouTube performance: Not yet available"}

${hasDir ? "Director Pass: Completed" : "Director Pass: Not run"}

In exactly 2 short sentences each, answer:
WORKED: What worked well that should be repeated in future videos?
IMPROVE: What should be improved or done differently next time?

Return ONLY:
WORKED: [sentence]
IMPROVE: [sentence]
`;

    const payload = {
      model      : ANTHROPIC_MODEL,
      max_tokens : 200,
      system     : [{ type: "text", text: "You are a concise video performance analyst for GovernX YouTube channel. Return only the two requested fields.", cache_control: { type: "ephemeral" } }],
      thinking   : { type: "adaptive" },
      output_config: { effort: "low" },
      messages   : [{ role: "user", content: prompt }]
    };

    const response = UrlFetchApp.fetch(ANTHROPIC_API_URL, {
      method            : "post",
      contentType       : "application/json",
      headers: {
        "x-api-key"         : apiKey,
        "anthropic-version" : "2023-06-01"
      },
      payload           : JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) throw new Error("API error");

    const json = JSON.parse(response.getContentText());
    const text = json.content.filter(b => b.type === "text").map(b => b.text).join("");

    const workedMatch  = text.match(/WORKED:\s*(.+)/);
    const improveMatch = text.match(/IMPROVE:\s*(.+)/);

    return {
      worked : workedMatch  ? workedMatch[1].trim()  : "Data captured — see QA scores",
      improve: improveMatch ? improveMatch[1].trim() : "Check pacing and hook scores above"
    };

  } catch (err) {
    Logger.log("Lessons learned generation failed (non-fatal): " + err.message);
    return {
      worked : hasQA ? "QA score: " + record.qaScores.totalScore + "/60 — " + record.qaScores.verdict : "No QA data",
      improve: hasPerf ? "CTR: " + record.performance.ctr + "% | Retention: " + record.performance.retention + "%" : "Publish video to capture performance data"
    };
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// READ QA SCORES FOR CONTENT ID
// Reads QA Critic scores saved to the Script Bank NOTE column
// After Stage 3 runs with v2.0, save scores to NOTE column for memory retrieval
// ══════════════════════════════════════════════════════════════════════════════
function readQAScoresForContent(ss, contentId) {
  try {
    const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
    if (!scriptSheet) return null;

    const scriptData = scriptSheet.getDataRange().getValues();
    for (let i = 1; i < scriptData.length; i++) {
      if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() !== contentId) continue;

      const note = scriptData[i][COL_SCRIPT.NOTE - 1].toString();
      if (!note || !note.includes("QA_SCORES:")) return null;

      // Parse QA scores from NOTE column
      // Format: QA_SCORES: hook=8|reverse=7|grc=9|arabic=8|pacing=7|factual=9|total=48|verdict=APPROVED
      const match = note.match(/QA_SCORES:\s*(.+)/);
      if (!match) return null;

      const parts = {};
      match[1].split("|").forEach(p => {
        const [k, v] = p.split("=");
        if (k && v) parts[k.trim()] = isNaN(v) ? v.trim() : parseInt(v);
      });

      return {
        hookScore             : parts.hook             || 0,
        reverseScore          : parts.reverse          || 0,
        grcScore              : parts.grc              || 0,
        arabicScore           : parts.arabic           || null,
        pacingScore           : parts.pacing           || 0,
        factualScore          : parts.factual          || 0,
        dataCompletenessScore : parts.dataCompleteness || 0,
        totalScore            : parts.total            || 0,
        verdict               : parts.verdict          || "UNKNOWN"
      };
    }
  } catch (err) {
    Logger.log("QA score read error: " + err.message);
  }
  return null;
}


// ══════════════════════════════════════════════════════════════════════════════
// SAVE QA SCORES TO SCRIPT BANK
// Called from generateScript() in Pipeline.gs after QA evaluation completes
// Saves scores to the NOTE column so Channel Memory can retrieve them later
// ══════════════════════════════════════════════════════════════════════════════
function saveQAScoresToScript(contentId, evalResult) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
    if (!scriptSheet) return;

    const scriptData = scriptSheet.getDataRange().getValues();
    for (let i = 1; i < scriptData.length; i++) {
      if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() !== contentId) continue;

      const scoreString =
        "QA_SCORES: " +
        "hook="            + (evalResult.hookScore             || 0) + "|" +
        "reverse="         + (evalResult.reverseScore          || 0) + "|" +
        "grc="             + (evalResult.grcScore              || 0) + "|" +
        "arabic="          + (evalResult.arabicScore           !== null ? evalResult.arabicScore : "NA") + "|" +
        "pacing="          + (evalResult.pacingScore           || 0) + "|" +
        "factual="         + (evalResult.factualScore          || 0) + "|" +
        "dataCompleteness="+ (evalResult.dataCompletenessScore || 0) + "|" +
        "total="           + (evalResult.totalScore            || 0) + "|" +
        "verdict="         + (evalResult.verdict               || "UNKNOWN");

      scriptSheet.getRange(i + 1, COL_SCRIPT.NOTE).setValue(scoreString);
      Logger.log("QA scores saved for: " + contentId + " → " + scoreString);
      return;
    }
  } catch (err) {
    Logger.log("QA score save error (non-fatal): " + err.message);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// VIEW CHANNEL MEMORY
// Menu action — opens the memory file in Drive for inspection
// ══════════════════════════════════════════════════════════════════════════════
function viewChannelMemory() {
  const ui = SpreadsheetApp.getUi();

  try {
    const file = getMemoryFile(false);
    if (!file) {
      ui.alert(
        "No Channel Memory Yet",
        "Channel Memory is empty.\n\n" +
        "Run Stage 10 on a completed video, then run '🧠 Write Channel Memory' " +
        "to create the first memory record.",
        ui.ButtonSet.OK
      );
      return;
    }

    const content = file.getBlob().getDataAsString();
    const records = JSON.parse(content || "[]");

    if (records.length === 0) {
      ui.alert("Channel Memory is empty.", "No records found.", ui.ButtonSet.OK);
      return;
    }

    let summary = "📊 CHANNEL MEMORY — " + records.length + " record(s)\n\n";
    records.slice(-5).reverse().forEach((rec, idx) => {
      summary += (idx + 1) + ". " + rec.contentId + " — " + rec.company + "\n";
      summary += "   Domain: " + rec.domain + " | " + rec.discipline + "\n";
      if (rec.qaScores && rec.qaScores.totalScore) {
        summary += "   QA: " + rec.qaScores.totalScore + "/60 — " + rec.qaScores.verdict + "\n";
      }
      if (rec.performance && rec.performance.hasData) {
        summary += "   Views: " + rec.performance.views +
                   " | CTR: " + rec.performance.ctr + "%" +
                   " | Retention: " + rec.performance.retention + "%\n";
      }
      summary += "\n";
    });

    summary += "Full memory file: " + file.getUrl();

    ui.alert("Channel Memory", summary, ui.ButtonSet.OK);

  } catch (err) {
    ui.alert("Memory Error", err.message, ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// CLEAR CHANNEL MEMORY
// Destructive — prompts confirmation before clearing
// ══════════════════════════════════════════════════════════════════════════════
function clearChannelMemory() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    "⚠️ Clear Channel Memory",
    "This will permanently delete all Channel Memory records.\n\n" +
    "This cannot be undone. Are you sure?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  try {
    const file = getMemoryFile(false);
    if (!file) {
      ui.alert("Nothing to clear — memory file does not exist.", "", ui.ButtonSet.OK);
      return;
    }

    file.setContent("[]");
    ui.alert("✅ Channel Memory cleared.", "", ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("Error clearing memory: " + err.message, "", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// GET MEMORY FILE
// Returns the Google Drive file object for the memory JSON
// createIfMissing = true → creates the file if it doesn't exist
// createIfMissing = false → returns null if missing
// ══════════════════════════════════════════════════════════════════════════════
function getMemoryFile(createIfMissing) {
  const rootFolder = getOrCreateRootProductionFolder();

  // Check if file already exists in the folder
  const files = rootFolder.getFilesByName(MEMORY_FILE_NAME);
  if (files.hasNext()) return files.next();

  // Not found
  if (!createIfMissing) return null;

  // Create new empty memory file
  const newFile = rootFolder.createFile(MEMORY_FILE_NAME, "[]", MimeType.PLAIN_TEXT);
  Logger.log("Channel Memory file created: " + newFile.getUrl());
  return newFile;
}


// ── Helper: get or create root production folder ──────────────────────────────
function getOrCreateRootProductionFolder() {
  try {
    return DriveApp.getFolderById("1yErZa4vpGB-iAqetCgXKWepwQn1tMDSQ");
  } catch(e) {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    return folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// PIPELINE.GS INTEGRATION
//
// ── 1. Inject memory into buildMasterContentPrompt() ─────────────────────────
// In Pipeline.gs, inside buildMasterContentPrompt(idea), add at the TOP
// of the return template string, just after the opening backtick:
//
//   ${readChannelMemory()}
//
// Full context — replace the opening of the return:
//
//   function buildMasterContentPrompt(idea) {
//     return `
//   ${readChannelMemory()}
//   You are generating the Master Content Table entry...
//
// ── 2. Save QA scores after Stage 3 evaluation ────────────────────────────────
// In Pipeline.gs, in generateScript(), after showEvalResult() returns proceed,
// add one line BEFORE writeScript():
//
//   if (proceed) saveQAScoresToScript(idea.id, evalResult);
//
// Full context:
//   const proceed = showEvalResult(evalResult, idea);
//   if (!proceed) { ... return; }
//   saveQAScoresToScript(idea.id, evalResult);   ← ADD THIS
//   writeScript(idea.id, raw, targetFormat, master, idea);
//
// ── 3. Menu additions (Menu.gs) ───────────────────────────────────────────────
// Add after Stage 10 in onOpen():
//
//   .addSeparator()
//   .addItem("🧠 Write Channel Memory",  "writeChannelMemoryForSelected")
//   .addItem("📊 View Channel Memory",   "viewChannelMemory")
//   .addItem("🗑️  Clear Channel Memory",  "clearChannelMemory")
// ══════════════════════════════════════════════════════════════════════════════