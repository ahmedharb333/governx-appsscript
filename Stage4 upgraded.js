// ════════════════════════════════════════════════════════════════════════════════
// STAGE 4 — Generate Visual Library (v3 — Director-Embedded, Two-Pass)
//
// ARCHITECTURE:
//   Pass 1 — Scene Generation:  Claude generates scene list from script
//   Pass 2 — Director Review:   Same session, Claude self-critiques every scene
//                                and fills REMOTION_DATA, REMOTION_STYLE,
//                                VOICEOVER_SYNC, SCENE_SCORE for non-AI scenes
//
// RESULT: Stage 4.5 (Director Pass) is no longer needed.
//         Visual Library is production-ready after Stage 4 alone.
// ════════════════════════════════════════════════════════════════════════════════

// ── Helper: read script data for a content ID ────────────────────────────────
function readScriptData_(ideaId) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let voiceover = "", narrative = "", sceneBlueprint = "", dataMoments = "";
  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() !== ideaId) continue;
    voiceover = scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1];
    narrative = scriptData[i][COL_SCRIPT.NARRATIVE_FLOW   - 1];
    const note = scriptData[i][COL_SCRIPT.NOTE - 1].toString();
    const sbMatch = note.match(/SCENE_BLUEPRINT:\n([\s\S]+?)(?:\n\nQA_SCORES:|$)/);
    if (sbMatch) sceneBlueprint = sbMatch[1].trim();
    const dmMatch = note.match(/DATA_MOMENTS:\n([\s\S]+?)(?:\n\nSCENE_BLUEPRINT:|$)/);
    if (dmMatch) dataMoments = dmMatch[1].trim();
    break;
  }
  return { voiceover, narrative, sceneBlueprint, dataMoments };
}


// ════════════════════════════════════════════════════════════════════════════════
// STAGE 4 — Generate Scenes (Pass 1)
// Generates scene list from SCENE_BLUEPRINT and writes to Visual Library
// Run Stage 4B after this to fill REMOTION_DATA / REMOTION_STYLE
// ════════════════════════════════════════════════════════════════════════════════
function generateScenes() {
  const idea = getActiveIdeaRow();
  if (!idea) return;
  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;

  const master = getMasterContent(idea.id);
  const ui     = SpreadsheetApp.getUi();
  const { voiceover, narrative, sceneBlueprint, dataMoments } = readScriptData_(idea.id);

  ui.alert(
    "Stage 4 — Generating Scene List",
    "Claude is building the scene list for: " + idea.company +
    "\n\nEstimated time: 20–40 seconds.",
    ui.ButtonSet.OK
  );

  try {
    const pass1Prompt = buildScenesPrompt(idea, master, voiceover, narrative, dataMoments, sceneBlueprint);
    const pass1Raw    = callClaude(pass1Prompt, "stage_4_scenes");

    const validation = validateOutput("VISUAL", pass1Raw, idea.id, null);
    const proceed    = showValidationResult("VISUAL", validation, idea.id);
    if (!proceed) {
      // Diagnostic: capture what Claude actually returned so a 0-scenes gate
      // failure can be traced (empty vs truncated vs wrong-format).
      const diag = "len=" + pass1Raw.length + " hasSCENE=" + /SCENE_\d+_START/.test(pass1Raw) +
                   " | head: " + pass1Raw.slice(0, 280).replace(/\s+/g, " ") +
                   " | tail: " + pass1Raw.slice(-160).replace(/\s+/g, " ");
      logError("Stage 4 — Scenes", idea.id, "Gate raw-response", diag);
      logError("Stage 4 — Scenes", idea.id, "Quality Gate Failed", validation.failures.join(" | "));
      return;
    }

    writeScenes(idea.id, pass1Raw);
    updatePipelineStatus_(idea.id, "S4", "✅");

    ui.alert(
      "✅ Stage 4 Complete — Scenes Written",
      "Visual Library filled for: " + idea.id + "\n\n" +
      "⚡ NEXT: Run Stage 4B — Director Review\n" +
      "This fills REMOTION_DATA, REMOTION_STYLE, VOICEOVER_SYNC, SCENE_SCORE\n" +
      "for every scene so Stage 8D can render correctly.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S4", "❌");
    logError("Stage 4 — Scenes", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 4 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}


// ════════════════════════════════════════════════════════════════════════════════
// STAGE 4B — Director Review (Pass 2)
// Reads saved scene list + script, fills cols 19–22 with structured Remotion data
// Run this after Stage 4 completes
// ════════════════════════════════════════════════════════════════════════════════
function runDirectorReview() {
  const idea = getActiveIdeaRow();
  if (!idea) return;
  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const master = getMasterContent(idea.id);
  const ui     = SpreadsheetApp.getUi();
  const { voiceover, sceneBlueprint } = readScriptData_(idea.id);

  // Read the Pass 1 scene rows to reconstruct the scene manifest
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();

  // Collected per scene so the review can be sent in BATCHES. One call for all
  // scenes worked at 11 scenes and stopped working at 27: the model must emit a
  // full REVIEW block per scene, and 27 of them plus adaptive thinking runs past
  // the ~6 minutes UrlFetchApp will wait — all three retries timed out. Smaller
  // batches keep each response well inside the limit, and a failed batch can be
  // retried on its own instead of losing the whole review.
  const sceneBlocks = [];
  let pass1Raw = "";
  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const sceneNum   = visualData[i][COL_VISUAL.SCENE_NUM    - 1];
    const sceneType  = visualData[i][COL_VISUAL.SCENE_TYPE   - 1];
    const desc       = visualData[i][COL_VISUAL.DESCRIPTION  - 1];
    const source     = visualData[i][COL_VISUAL.SOURCE       - 1];
    const cpDate     = visualData[i][COL_VISUAL.CHECKPOINT_DATE  - 1];
    const cpEvent    = visualData[i][COL_VISUAL.CHECKPOINT_EVENT - 1];
    const cpAngle    = visualData[i][COL_VISUAL.CHECKPOINT_ANGLE - 1];
    const timestamp  = visualData[i][COL_VISUAL.TIMESTAMP    - 1];
    const block =
      "SCENE_" + sceneNum + "_START\n" +
      "SCENE_NUM: "       + sceneNum  + "\n" +
      "SCENE_TYPE: "      + sceneType + "\n" +
      "DESCRIPTION: "     + desc      + "\n" +
      "SCENE_SOURCE: "    + source    + "\n" +
      "SCENE_TIMESTAMP: " + timestamp + "\n" +
      (cpDate  ? "CHECKPOINT_DATE: "  + cpDate  + "\n" : "") +
      (cpEvent ? "CHECKPOINT_EVENT: " + cpEvent + "\n" : "") +
      (cpAngle ? "CHECKPOINT_ANGLE: " + cpAngle + "\n" : "") +
      "SCENE_" + sceneNum + "_END\n\n";
    sceneBlocks.push({ num: sceneNum, text: block });
    pass1Raw += block;
  }

  if (!pass1Raw) {
    ui.alert("No scenes found for: " + idea.id + "\nRun Stage 4 first.");
    return;
  }

  // ~8 scenes per call keeps each response inside the UrlFetchApp timeout.
  const BATCH = 8;
  const batches = [];
  for (let i = 0; i < sceneBlocks.length; i += BATCH) batches.push(sceneBlocks.slice(i, i + BATCH));

  ui.alert(
    "Stage 4B — Director Review",
    "Running Director review for: " + idea.company + "\n\n" +
    sceneBlocks.length + " scenes in " + batches.length + " batch(es) of up to " + BATCH + ".\n" +
    "Fills REMOTION_DATA, REMOTION_STYLE, VOICEOVER_SYNC, SCENE_SCORE.\n\n" +
    "Estimated time: about " + (batches.length * 2) + "–" + (batches.length * 4) + " minutes.\n" +
    "Each batch is written as it completes, so nothing is lost if a later one fails.",
    ui.ButtonSet.OK
  );

  try {
    let pass2Raw = "";
    let done = 0;

    for (let b = 0; b < batches.length; b++) {
      const chunk = batches[b];
      const chunkRaw = chunk.map(function (s) { return s.text; }).join("");

      // SCENE_REF is resolved by INDEX on write, so a batch must return the
      // absolute scene number — a per-batch counter would write batch 2's data
      // onto scenes 1-8. Stated explicitly rather than relying on inference.
      const batchNote =
        "\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "THIS IS BATCH " + (b + 1) + " OF " + batches.length + "\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "Review ONLY the " + chunk.length + " scenes listed above " +
        "(scene numbers " + chunk[0].num + " to " + chunk[chunk.length - 1].num + ").\n" +
        "⚠ SCENE_REF MUST BE THE SCENE_NUM SHOWN — not a counter starting at 1.\n" +
        "   For the first scene in this batch, SCENE_REF is " + chunk[0].num + ".\n" +
        "Return exactly " + chunk.length + " REVIEW blocks.";

      const prompt = buildDirectorReviewPrompt(idea, master, voiceover, chunkRaw, sceneBlueprint || "") + batchNote;
      const raw    = callClaude(prompt, "stage_4_director");

      // Write each batch as it lands — a later failure cannot undo earlier work.
      writeDirectorReviewResults(idea.id, raw);
      SpreadsheetApp.flush();

      pass2Raw += raw;
      done += (raw.match(/REVIEW_\d+_START/g) || []).length;
      Logger.log("Stage 4B: batch " + (b + 1) + "/" + batches.length +
                 " → " + (raw.match(/REVIEW_\d+_START/g) || []).length + " reviews");
    }

    if (done < sceneBlocks.length) {
      ui.alert("⚠️ Partial review",
        "Reviewed " + done + " of " + sceneBlocks.length + " scenes.\n\n" +
        "Re-run Stage 4B to fill the rest — scenes already written are kept.",
        ui.ButtonSet.OK);
    }

    const scoreMatches = pass2Raw.match(/SCENE_SCORE:\s*([^\n]+)/g) || [];
    const aCount       = scoreMatches.filter(s => s.includes(": A")).length;
    const bCount       = scoreMatches.filter(s => s.includes(": B")).length;
    const needsData    = scoreMatches.filter(s => s.includes("NEEDS_DATA")).length;
    const templateReq  = scoreMatches.filter(s => s.includes("TEMPLATE_REQUEST")).length;

    ui.alert(
      "✅ Stage 4B Complete — Director Review Done",
      "REMOTION_DATA filled for: " + idea.id + "\n\n" +
      "SCENE QUALITY SCORES:\n" +
      "  ✅ A (strong)        : " + aCount       + "\n" +
      "  ⚠️  B (acceptable)   : " + bCount       + "\n" +
      "  ❌ NEEDS_DATA        : " + needsData    + "\n" +
      "  🔧 TEMPLATE_REQUEST : " + templateReq  + "\n\n" +
      (needsData  > 0 ? "⚠️  " + needsData  + " scene(s) need source data.\n" : "") +
      (templateReq > 0 ? "🔧 " + templateReq + " scene(s) need new templates.\n" : "") +
      "\nProceed to Stage 5 — Create Publishing Row.",
      ui.ButtonSet.OK
    );

  } catch (err) {
    logError("Stage 4B — Director", idea.id, "API / Runtime Error", err.message);
    ui.alert("❌ Stage 4B Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}


// ════════════════════════════════════════════════════════════════════════════════
// PASS 2 PROMPT — Director Review + Remotion Data Writer
// Sees the full voiceover + the Pass 1 scene list
// Outputs REMOTION_DATA, REMOTION_STYLE, VOICEOVER_SYNC, SCENE_SCORE per scene
// ════════════════════════════════════════════════════════════════════════════════
function buildDirectorReviewPrompt(idea, master, voiceover, pass1Raw, sceneBlueprint) {

  // Same evidence guardrail as Pass 1: REMOTION_DATA is what actually renders,
  // so the verified whitelist must gate this pass too. (Helper is in pipeline.gs.)
  const verifiedFiguresBlock =
    (typeof buildVerifiedFiguresBlock_ === "function") ? buildVerifiedFiguresBlock_(idea.id) : "";

  // Extract scene blocks from Pass 1 for review
  const sceneBlocks = pass1Raw.match(/SCENE_\d+_START[\s\S]*?SCENE_\d+_END/g) || [];
  const sceneManifest = sceneBlocks.map((block, i) => {
    const get = (f) => { const m = block.match(new RegExp(f + ":\\s*([^\\n\\r]*)")); return m ? m[1].trim() : ""; };
    const num  = i + 1;
    const type = get("SCENE_TYPE");
    const desc = get("DESCRIPTION");
    const src  = get("SCENE_SOURCE");
    const cpDate  = get("CHECKPOINT_DATE");
    const cpEvent = get("CHECKPOINT_EVENT");
    const cpAngle = get("CHECKPOINT_ANGLE");
    const ts   = get("SCENE_TIMESTAMP");

    let entry = `SCENE_${num} | ${ts} | ${type.toUpperCase()} | Source: ${src}`;
    if (type === "Checkpoint") {
      entry += `\n  Date: ${cpDate} | Event: ${cpEvent} | Angle: ${cpAngle}`;
    } else {
      entry += `\n  Visual: ${desc.substring(0, 120)}`;
    }
    return entry;
  }).join("\n\n");

  return `
You are the Creative Director for GovernX. You have just received a completed scene list.
Your job now is to review every scene, judge its quality, and write the structured data
that Stage 8D (Remotion renderer) will read to build each scene.

This is a production instruction — not a review essay.
Every output you write will be parsed by code and written directly to the Visual Library.
Write precisely, in the exact format specified.
${verifiedFiguresBlock}
═══════════════════════════════════════════════════════
CONTENT BRIEF  (⚠ may contain unverified figures — the block above overrides it)
═══════════════════════════════════════════════════════
Company       : ${idea.company}
Title         : ${master ? master.title : idea.company}
Discipline    : ${master ? master.discipline : "GRC"}
Core Insight  : ${master ? master.coreInsight : ""}
Checkpoints   : ${master ? master.checkpoints : ""}

FULL VOICEOVER SCRIPT:
${voiceover}

SCENE BLUEPRINT FROM STAGE 3 (USE THIS AS YOUR DATA SOURCE):
${sceneBlueprint || "Not available — infer from voiceover and scene list"}

CRITICAL: The SCENE_BLUEPRINT above contains the exact data for every scene.
When filling REMOTION_DATA, read the DISPLAY field from the matching SCENE_BP entry.
Do NOT invent data. Do NOT approximate. Copy the exact figures, labels, and years
that Stage 3 already identified from the research sources.

═══════════════════════════════════════════════════════
SCENE LIST FROM PASS 1
═══════════════════════════════════════════════════════
${sceneManifest}

═══════════════════════════════════════════════════════
YOUR THREE TASKS PER SCENE
═══════════════════════════════════════════════════════

For EVERY scene, you must output a REVIEW block in the format below.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TASK 1 — VOICEOVER_SYNC
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Copy the EXACT sentence(s) from the voiceover that this scene plays under.
Do not paraphrase. Do not summarize. Copy the words verbatim.
If a scene covers multiple sentences, include all of them.
For AI Generated / Stock scenes: this is the B-roll coverage window.
For Checkpoint scenes: this is the moment the voiceover names this checkpoint.
For Text / Infographic scenes: this is the data sentence this scene makes visible.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TASK 2 — SCENE_SCORE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Score each scene honestly. Choose ONE of these four values:

A           — Strong. Visual earns its voiceover beat. Data is shown when it must be shown.
              Checkpoint has clean date/event/angle. No disconnect between word and image.

B           — Acceptable but improvable. Minor mismatch or generic choice that works but
              does not excel. Note WHY it is B in the SCORE_REASON field.

NEEDS_DATA  — The scene needs quantitative data that is not available in the script.
              This blocks REMOTION_DATA from being filled correctly.
              Specify what data is needed in SCORE_REASON.

TEMPLATE_REQUEST — The ideal visual for this scene does not fit any existing Remotion template.
              Describe the needed template in SCORE_REASON.
              Still fill REMOTION_DATA and REMOTION_STYLE with the closest existing template.

DIRECTOR SCORING CRITERIA:
A scene scores A when ALL of these are true:
  ✓ The visual would communicate the story beat even with sound off
  ✓ Data moments (any stat, comparison, trend) are shown as Text or Infographic — never B-roll
  ✓ AI Generated scenes use atmospheric metaphor, not literal illustration
  ✓ Checkpoint cards have precise date, sharp event sentence, and governance angle
  ✓ The emotional weight of the voiceover moment is matched by the visual intensity

A scene scores B when ONE of these is true:
  • Generic B-roll covers a specific data moment (voiceover says "$83B" but scene is a boardroom)
  • The scene is correct but the Infographic type is suboptimal
  • The KlingAI prompt is atmospheric but not specific enough to the story beat

A scene scores NEEDS_DATA when:
  • The voiceover references a specific statistic that is not cited anywhere in the script
  • You cannot write a complete REMOTION_DATA block because a key number is missing

A scene scores TEMPLATE_REQUEST when:
  • The ideal visual requires animation logic not possible in existing templates
  • Describe exactly what the new template needs to do

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
TASK 3 — REMOTION_DATA and REMOTION_STYLE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
These two fields are ONLY required for scenes that Remotion will render:
  - Checkpoint scenes
  - Text scenes  
  - Infographic scenes

For Timeline scenes:
   REMOTION_DATA: checkpoints=[comma-separated list of all checkpoint dates in chronological order]
   REMOTION_STYLE: template=Timeline | variant=standard | weight=closing
   Example: REMOTION_DATA: checkpoints=IPO 2012,Debt Concealment 2012-2019,EY Audit Failure 2013-2019,Collapse February 2020
   This is NOT N/A — the Timeline needs the checkpoint list to render.

For all other scene types (Text, Infographic, Data Table, Checkpoint): fill both fields as specified above.

REMOTION_DATA FORMAT:
Write as pipe-separated key=value pairs. No quotes. No JSON. No brackets.
Use only the keys defined for each template below.

REMOTION_STYLE FORMAT:
Write as pipe-separated key=value pairs.
Always include: template=X | variant=Y | weight=normal|root|outcome

═══════════════════════════════════════════════════════
REMOTION TEMPLATE REFERENCE
═══════════════════════════════════════════════════════

TEMPLATE: CheckpointCard
  Used for: CHECKPOINT scenes
  REMOTION_DATA keys:
    date=       [year or short date, e.g. "September 2016" or "2002–2007"]
    event=      [what happened — sharp, max 12 words]
    angle=      [governance/process angle in CAPS — max 10 words]
  REMOTION_STYLE:
    template=CheckpointCard | variant=standard|root|outcome | weight=normal|root
    Use variant=root for ROOT CAUSE checkpoint. Use variant=outcome for first checkpoint.
  Example:
    REMOTION_DATA: date=January 2013 | event=BB10 launches three years behind schedule | angle=PROCESS FAILURE — NO STAGE-GATE DELIVERY DISCIPLINE
    REMOTION_STYLE: template=CheckpointCard | variant=standard | weight=normal

TEMPLATE: TextImpactScene
  Used for: TEXT scenes (hook cards, verdict cards, quote cards)
  REMOTION_DATA keys:
    mainText=   [primary text to display — quoted values extracted from description]
    subText=    [secondary line, if any — leave blank if none]
    type=       [default | shatter | verdict]
                shatter: first value shatters, second value reveals (for dramatic stats)
                verdict: "NOT X. A Y." format (for closing argument cards)
                default: standard sequential punch-in
  REMOTION_STYLE:
    template=TextImpactScene | variant=default|shatter|verdict | weight=normal|hook|closing
  Example (shatter type):
    REMOTION_DATA: mainText=$83 BILLION | subText=50% U.S. MARKET SHARE | type=shatter
    REMOTION_STYLE: template=TextImpactScene | variant=shatter | weight=hook

TEMPLATE: InfographicScene
  Used for: INFOGRAPHIC scenes
  Sub-templates (variant):

  variant=LINE_GRAPH
    REMOTION_DATA keys:
      type=LINE_GRAPH
      label=    [chart title, e.g. "U.S. Smartphone Market Share"]
      unit=     [% | $B | $M | units — whatever the y-axis represents]
      points=   [comma-separated year:value pairs, e.g. "2009:50,2010:38,2013:3,2016:1"]
      highlight=[which data point to emphasize, e.g. "2016:1"]
      duration= [seconds, e.g. 4]
    Example:
      REMOTION_DATA: type=LINE_GRAPH | label=BlackBerry U.S. Market Share | unit=% | points=2009:50,2010:38,2011:24,2012:10,2013:3,2016:1 | highlight=2016:1 | duration=4
      REMOTION_STYLE: template=InfographicScene | variant=LINE_GRAPH | weight=normal

  variant=SPLIT_COMPARISON
    REMOTION_DATA keys:
      type=SPLIT_COMPARISON
      left_label=   [left side name]
      left_values=  [comma-separated label:value pairs for left side]
      right_label=  [right side name]
      right_values= [comma-separated label:value pairs for right side]
      bottom_note=  [optional callout text below the comparison]
      duration=     [seconds]
    Example:
      REMOTION_DATA: type=SPLIT_COMPARISON | left_label=Apple App Store | left_values=2008:500,2010:300000 | right_label=BlackBerry App World | right_values=2009:1000,2010:20000 | bottom_note=2008–2010 App Ecosystem Gap | duration=4
      REMOTION_STYLE: template=InfographicScene | variant=SPLIT_COMPARISON | weight=normal

  variant=DATA_CALLOUT
    REMOTION_DATA keys:
      type=DATA_CALLOUT
      value=        [the primary number/stat, e.g. "$77.5M" or "50%"]
      label=        [what the value represents]
      context=      [one sentence explaining significance — 10 words max]
      duration=     [seconds]
    Example:
      REMOTION_DATA: type=DATA_CALLOUT | value=$77.5M | label=Options Backdating Settlement | context=Board oversight failure at governance crisis moment | duration=4
      REMOTION_STYLE: template=InfographicScene | variant=DATA_CALLOUT | weight=normal

  variant=COUNTER_ANIMATION
    REMOTION_DATA keys:
      type=COUNTER_ANIMATION
      from=     [starting value, e.g. "0"]
      to=       [ending value, e.g. "$83B" or "300,000"]
      unit=     [prefix/suffix, e.g. "$" or "%" or "apps"]
      label=    [what is being counted]
      duration= [seconds]
    Example:
      REMOTION_DATA: type=COUNTER_ANIMATION | from=0 | to=83 | unit=$B | label=BlackBerry Peak Market Cap | duration=3

  variant=BEFORE_AFTER_CARD
    REMOTION_DATA keys:
      type=BEFORE_AFTER_CARD
      before_label=   [left column header]
      after_label=    [right column header]
      before_rows=    [rows separated by ";;" — each row is "item → value"]
      after_rows=     [rows separated by ";;" — each row is "item → value"]
      verdict=        [optional red bottom line]
      duration=       [seconds]
    ⚠ ROWS USE ";;" — NEVER "|". A "|" starts a NEW KEY and destroys the block.
    Example:
      REMOTION_DATA: type=BEFORE_AFTER_CARD | before_label=Governance as designed | before_rows=Pay setting → Committee;;Oversight → Independent board;;Disclosure → Full | after_label=Nissan 2004-2018 | after_rows=Pay setting → Chairman alone;;Oversight → None;;Disclosure → Split across years | verdict=ONE DELEGATION REMOVED EVERY CHECK | duration=12
      REMOTION_STYLE: template=InfographicScene | variant=BEFORE_AFTER_CARD | weight=normal

  ─────────────────────────────────────────────────────
  THE FOUR UNDER-USED VARIANTS. All four render today. A film that never reaches
  for them looks like the same two cards repeating — use them wherever the
  verified evidence supports one.
  ─────────────────────────────────────────────────────

  variant=BAR_CHART            → comparing magnitudes across 3–6 items
    REMOTION_DATA keys:
      type=BAR_CHART
      title=      [what is being compared]
      unit=       [$M | $B | % | count]
      points=     [items separated by "," — each "Label:value:highlight"]
    Example:
      REMOTION_DATA: type=BAR_CHART | title=Penalties and settlements | unit=$M | points=SEC penalty - Nissan:15:true,Ghosn penalty:1:false,Kelly penalty:0.1:false | duration=12
      REMOTION_STYLE: template=InfographicScene | variant=BAR_CHART | weight=normal

  variant=KPI_DASHBOARD        → 3–4 headline figures at once, each attributed
    REMOTION_DATA keys:
      type=KPI_DASHBOARD
      title=      [e.g. "The case in numbers"]
      kpis=       [items separated by "," — each
                   "label:value:trend:change:context:highlight:tag"
                   trend and change may be left EMPTY; tag = Regulator|Company|Court]
    ⚠ No ":" or "," inside a label or context — use "-" instead.
    Example:
      REMOTION_DATA: type=KPI_DASHBOARD | title=The case in numbers | kpis=Concealed pay:$140M:::FY2009-2018:true:Regulator,Nissan penalty:$15M:::SEC settlement:false:Regulator,Officer bar:10 years:::Ghosn:false:Regulator | duration=15
      REMOTION_STYLE: template=InfographicScene | variant=KPI_DASHBOARD | weight=normal

  variant=RISK_MATRIX          → where the control gap lived
    REMOTION_DATA keys:
      type=RISK_MATRIX
      title=      [e.g. "Where the control gap lived"]
      risks=      [items separated by "," — each
                   "label:likelihood:impact:highlight"   (1=low 2=med 3=high)]
    Example:
      REMOTION_DATA: type=RISK_MATRIX | title=Where the control gap lived | risks=Compensation self-setting:3:3:true,Board oversight:2:3:false,Disclosure review:2:2:false | duration=14
      REMOTION_STYLE: template=InfographicScene | variant=RISK_MATRIX | weight=normal

  variant=PROGRESS_GAUGE       → coverage / completeness as a percentage
    REMOTION_DATA keys:
      type=PROGRESS_GAUGE
      title=      [e.g. "Governance coverage"]
      gauges=     [items separated by "," — each
                   "label:value:unit:context:highlight"]
    Example:
      REMOTION_DATA: type=PROGRESS_GAUGE | title=Governance coverage | gauges=Independent directors:11:%:one of nine:true,Audit committee:0:%:none existed:true | duration=12
      REMOTION_STYLE: template=InfographicScene | variant=PROGRESS_GAUGE | weight=normal

═══════════════════════════════════════════════════════
CASE-FILE COMPONENTS — THE PREMIUM LOOKS
═══════════════════════════════════════════════════════
These eleven are NOT InfographicScene variants — they are dedicated case-file
components, and they are what makes a film look authored rather than templated.
They render today. Reach for them whenever the evidence supports one.

⚠ THREE ROUTING RULES — get these wrong and the scene renders as something else:
   • Set Scene Type = "Infographic" or "Text" for ALL of these.
     A Scene Type of "Timeline" or "Checkpoint" is intercepted BEFORE type= is
     read, so BEAT_TIMELINE with Scene Type=Timeline renders as a plain timeline.
   • type=VERDICT_CARD is the VerdictCard component.
     type=verdict (no _CARD) is a StatementCard. They are different scenes.
   • Never put an "angle=" or "checkpoints=" key in these — either one re-routes
     the scene to a Checkpoint or Timeline card.

  type=EVIDENCE_CARD      → THE TRUST UNIT. A verbatim quote from a source
                            document on a cream card. Use for the single
                            strongest piece of evidence in the film.
    keys: headline= value= value_label= extract= attribution=
          source_publisher= source_year= source_doc= verified=true|false
    ⚠ extract= must be a VERBATIM quote from the source. Never paraphrase.
    Example:
      REMOTION_DATA: type=EVIDENCE_CARD | headline=SEC found Nissan concealed compensation | value=$140M | value_label=CONCEALED PAY - FY2009-2018 | extract=Nissan failed to disclose more than $140 million in compensation and retirement benefits | attribution=Regulator | source_publisher=SEC | source_year=2019 | source_doc=Press release 2019-183 | verified=true | duration=14
      REMOTION_STYLE: template=EvidenceCard | variant=standard | weight=root

  type=DATA_WALL          → 3–4 headline figures, EACH with its own attribution
                            tag. The "case in numbers" beat.
    keys: title= rows=value:label:sourceType:highlight, … source_publisher=
          sourceType = Regulator | Company | Court
    Example:
      REMOTION_DATA: type=DATA_WALL | title=The case in numbers | rows=$140M:concealed compensation:Regulator:true,$15M:Nissan civil penalty:Regulator:false,10 years:officer and director bar:Regulator:false | source_publisher=SEC | duration=16
      REMOTION_STYLE: template=CaseDataWall | variant=standard | weight=normal

  type=VERDICT_CARD       → the GOVERNX VERDICT close. Use ONCE, as the
                            second-to-last scene.
    keys: ruling= punch= sign_off=
    Example:
      REMOTION_DATA: type=VERDICT_CARD | ruling=This was not one person acting alone. It was an architecture that made one person unchallengeable. | punch=Governance fails when no one can say no. | duration=12
      REMOTION_STYLE: template=VerdictCard | variant=standard | weight=outcome

  type=OPENING_HOOK       → fast-assembling scandal poster. Use ONCE, scene 1,
                            INSTEAD of a plain Text card, when there is one
                            dominant number to open on.
    keys: company= kicker= value= unit= label= secondary=
          source_publisher= source_year=
    Example:
      REMOTION_DATA: type=OPENING_HOOK | company=NISSAN | kicker=THE GHOSN GOVERNANCE CRISIS | value=140 | unit=$M | label=CONCEALED COMPENSATION | secondary=hidden across nine fiscal years | source_publisher=SEC | source_year=2019 | duration=10
      REMOTION_STYLE: template=OpeningHook | variant=standard | weight=normal

  type=STAT_POSTER        → ONE clean formatted number + caption + attribution.
                            Use when the figure is already formatted text
                            ("9.078 billion yen") rather than a count-up.
    keys: kicker= value= label= sublabel= attribution= source_publisher=
    Example:
      REMOTION_DATA: type=STAT_POSTER | value=9.078 billion yen | label=CONCEALED COMPENSATION | sublabel=FY2009-FY2017 | attribution=Company self-reported | source_publisher=Nissan | duration=9
      REMOTION_STYLE: template=StatPoster | variant=standard | weight=normal

  type=DECISION_CHAIN     → the causal chain: decision → consequence → failure.
    keys: title= nodes=A → B → C → D  outcome= source_publisher=
    Example:
      REMOTION_DATA: type=DECISION_CHAIN | title=The decision chain | nodes=Board delegates pay authority → No independent committee formed → Chairman sets own pay → Disclosure split across years | outcome=GOVERNANCE LESSON | source_publisher=SEC | duration=15
      REMOTION_STYLE: template=DecisionChain | variant=standard | weight=root

  type=CONTROL_GAP        → which control layer was present and which was
                            missing. The single clearest "where it broke" visual.
    keys: title= layers=Label:ok, Label:gap:why, …  outcome=
          use "ok" for a working layer, "gap" for the failed one
    Example:
      REMOTION_DATA: type=CONTROL_GAP | title=Where the control should have been | layers=Board oversight:ok,Internal audit:ok,Compensation setting:gap:no independent committee existed,Disclosure review:ok | outcome=$140M concealed over nine years | duration=15
      REMOTION_STYLE: template=ControlGapMap | variant=standard | weight=root

  type=CONTROL_PERIMETER  → what sat inside the oversight perimeter and the one
                            thing that sat outside it.
    keys: title= inside=A,B,C  outside= outside_note= source_publisher=
    Example:
      REMOTION_DATA: type=CONTROL_PERIMETER | title=Everything oversight covered - and the one thing it did not | inside=AUDIT,RISK COMMITTEE,COMPLIANCE | outside=CHAIRMAN COMPENSATION | outside_note=set by the chairman himself - never independently reviewed | duration=14
      REMOTION_STYLE: template=ControlPerimeter | variant=standard | weight=root

  type=CLAIM_LEDGER       → every figure in the film with its source, on one
                            card. Optional closing credibility beat.
    keys: title= rows=claim → source → id, …
    Example:
      REMOTION_DATA: type=CLAIM_LEDGER | title=Every figure on the record | rows=$140M concealed compensation → SEC 2019 → C1,$15M civil penalty → SEC 2019 → C11,9.078 billion yen → Nissan 2018 → C31 | duration=15
      REMOTION_STYLE: template=ClaimLedger | variant=standard | weight=normal

  type=BEAT_TIMELINE      → a compact THREE-beat timeline. Use mid-film when the
                            full timeline would be too dense.
                            (Scene Type must be Infographic, NOT Timeline.)
    keys: title= beats=YEAR → event,YEAR → event,YEAR → event
    Example:
      REMOTION_DATA: type=BEAT_TIMELINE | title=Three moments that decided it | beats=1999 → Revival Plan builds unchallengeable status,2004 → Board delegates compensation authority,2018 → Arrest at Haneda Airport | duration=13
      REMOTION_STYLE: template=BeatTimeline | variant=standard | weight=normal

  type=GOVERNANCE_METHOD  → the GovernX 4-step method applied to this case.
    keys: steps=TAG:title:detail, …  failure_index=  (0-based, which step failed)
          tags in order: OUTCOME, DECISION CHAIN, GOVERNANCE FAILURE, CONTROL LESSON
    Example:
      REMOTION_DATA: type=GOVERNANCE_METHOD | steps=OUTCOME:$140M concealed:over nine fiscal years,DECISION CHAIN:Board delegated pay authority:2004 with no committee,GOVERNANCE FAILURE:No one could challenge the chairman:for fourteen years,CONTROL LESSON:Pay authority needs an independent committee:always | failure_index=2 | duration=16
      REMOTION_STYLE: template=GovernanceMethod | variant=standard | weight=root

═══════════════════════════════════════════════════════
DIRECTOR QUALITY RULES (enforce on every scene)
═══════════════════════════════════════════════════════

0. PACING AND VARIETY — CHECK THESE ACROSS THE WHOLE FILM, NOT PER SCENE.
   A scene stays on screen exactly as long as the VOICEOVER_SYNC you assign it.
   That is the only timing control, so these are your responsibility:

   a) NO SCENE MAY CARRY MORE THAN 58 WORDS (~25 seconds) of VOICEOVER_SYNC.
      Over that, score it B and say in SCORE_REASON where to split it.
   b) A scene whose narration ENUMERATES A LIST must not be one scene.
      Six figures listed = six scenes.
   c) AT LEAST 8 DIFFERENT variants across the film; NO variant more than 4 times;
      Checkpoint cards MAXIMUM 4.
      AT LEAST 3 of those must be CASE-FILE COMPONENTS (EVIDENCE_CARD, DATA_WALL,
      CONTROL_GAP, DECISION_CHAIN, CONTROL_PERIMETER, OPENING_HOOK, VERDICT_CARD,
      STAT_POSTER, BEAT_TIMELINE, CLAIM_LEDGER, GOVERNANCE_METHOD). A film built
      only from Text / Checkpoint / Infographic looks templated — those eleven
      are what make it look authored.
      STRONGLY PREFERRED in every film:
        • ONE EVIDENCE_CARD — the verbatim-quote trust unit
        • ONE VERDICT_CARD  — as the second-to-last scene
        • ONE of CONTROL_GAP or DECISION_CHAIN — the "where it broke" analysis
   d) NEVER two scenes of the same variant back to back. If the scene list has
      them, re-assign one to a different variant that its data supports.
   e) Match dwell to reading load:
        5–8s  (12–19 words) — one statement, one big number
        8–12s (19–28 words) — checkpoint, single comparison
        12–18s (28–42 words) — KPI dashboard, risk matrix, bar chart, timeline

   Real failure this exists to prevent: a film gave ONE scene 103 seconds of
   narration listing six figures — one static poster for 30% of the runtime —
   and used Checkpoint for 7 of its 11 scenes. Both are scoring failures.

1. DATA MUST BE SHOWN — not just spoken.
   If the voiceover says a number, percentage, comparison, or trend,
   and there is no Text or Infographic scene covering it: the score is B or lower.

2. SPECIFICITY — no generic choices.
   A checkpoint ANGLE must name the exact governance/process failure type — not a vague label.
   An infographic LABEL must match the exact variable being measured (e.g. "U.S. Market Share %" not "Performance").
   A Text card mainText must be the actual quoted value from the voiceover, not a description of it.
   A COUNTER_ANIMATION must use the exact figures from the voiceover — never approximate.

3. VOICEOVER SYNC IS MANDATORY — AND NEVER DUPLICATED.
   Every Remotion scene must map to specific voiceover words.
   "The board watched" is not enough — copy the exact sentence from the script.
   ⚠ NO TWO SCENES MAY CARRY THE SAME VOICEOVER_SYNC. Each scene owns a DISTINCT,
   consecutive span of the script. If two scenes you are reviewing were assigned
   the same sentence, that is a Pass-1 padding error: score the weaker one B and
   say in SCORE_REASON that the scenes should be merged or one dropped — do not
   copy a sentence onto a second scene to fill it. The closing especially: the
   verdict and the CTA are ONE scene each, spoken once.

4. ROOT CAUSE SCENE gets variant=root.
   The final checkpoint (root cause) must use variant=root in REMOTION_STYLE.
   This triggers the red text treatment in CheckpointCard.

5. TEMPLATE_REQUEST means the request is LOGGED, not blocked.
   Still fill REMOTION_DATA and REMOTION_STYLE with the closest existing template.
   Write what the new template needs to do in SCORE_REASON.
   Rendering continues — the new template request is a future improvement.

═══════════════════════════════════════════════════════
OUTPUT FORMAT — ONE REVIEW BLOCK PER SCENE
═══════════════════════════════════════════════════════

REVIEW_1_START
SCENE_REF: 1
VOICEOVER_SYNC: [exact voiceover sentence(s) this scene covers]
SCENE_SCORE: [A | B | NEEDS_DATA | TEMPLATE_REQUEST]
SCORE_REASON: [required if B, NEEDS_DATA, or TEMPLATE_REQUEST — one sentence]
REMOTION_DATA: [structured key=value string OR N/A]
REMOTION_STYLE: [structured key=value string OR N/A]
REVIEW_1_END

REVIEW_2_START
[repeat for every scene]
REVIEW_N_END

${VISUAL_INTELLIGENCE_SKILL}
`;
}


// ════════════════════════════════════════════════════════════════════════════════
// WRITE PASS 2 RESULTS — fills cols 19–22 for each scene
// ════════════════════════════════════════════════════════════════════════════════
function writeDirectorReviewResults(contentId, pass2Raw) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();

  // Build a map of scene index → sheet row for this contentId
  // (Pass 1 wrote them in order, so index 0 = first scene row for this ID)
  const sceneRows = [];
  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() === contentId) {
      sceneRows.push(i + 1); // 1-based sheet row
    }
  }

  // Parse all REVIEW blocks from Pass 2
  const reviewBlocks = pass2Raw.match(/REVIEW_\d+_START([\s\S]*?)REVIEW_\d+_END/g) || [];

  reviewBlocks.forEach((block) => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*([^\\n\\r]*)"));
      return m ? m[1].trim() : "";
    };

    const sceneRef = parseInt(get("SCENE_REF"));
    if (isNaN(sceneRef) || sceneRef < 1) return;

    // sceneRef is 1-based scene index
    const sheetRow = sceneRows[sceneRef - 1];
    if (!sheetRow) return;

    const voiceoverSync  = get("VOICEOVER_SYNC");
    const sceneScore     = get("SCENE_SCORE");
    const scoreReason    = get("SCORE_REASON");
    const remotionData   = get("REMOTION_DATA");
    const remotionStyle  = get("REMOTION_STYLE");

    // Clear validation before writing
    [COL_VISUAL_EXTENDED.REMOTION_DATA,
     COL_VISUAL_EXTENDED.REMOTION_STYLE,
     COL_VISUAL_EXTENDED.VOICEOVER_SYNC,
     COL_VISUAL_EXTENDED.SCENE_SCORE].forEach(col => {
      visualSheet.getRange(sheetRow, col).clearDataValidations();
    });

    // Write to new columns (19–22)
    if (remotionData && remotionData !== "N/A") {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.REMOTION_DATA).setValue(remotionData).setWrap(true);
    }
    if (remotionStyle && remotionStyle !== "N/A") {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.REMOTION_STYLE).setValue(remotionStyle).setWrap(true);
    }
    if (voiceoverSync) {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.VOICEOVER_SYNC).setValue(voiceoverSync).setWrap(true);
    }

    // Scene score: include reason if present
    const scoreVal = sceneScore + (scoreReason ? " — " + scoreReason : "");
    if (sceneScore) {
      const scoreCell = visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.SCENE_SCORE);
      scoreCell.setValue(scoreVal).setWrap(true);

      // Color-code the score cell
      const bg = sceneScore === "A"                ? "#D4EDDA" :  // green
                 sceneScore === "B"                ? "#FFF3CD" :  // yellow
                 sceneScore === "NEEDS_DATA"       ? "#F8D7DA" :  // red
                 sceneScore === "TEMPLATE_REQUEST" ? "#D1ECF1" :  // blue
                 "#FFFFFF";
      scoreCell.setBackground(bg);
    }

    Logger.log("Stage 4 Pass 2 — Scene " + sceneRef + ": " + sceneScore);
  });

  SpreadsheetApp.flush();
  Logger.log("Stage 4 Pass 2 complete — " + reviewBlocks.length + " scenes reviewed.");

  // ── Duplicate-voiceover guard (code-enforced) ──────────────────────────────
  // The prompt forbids two scenes sharing a VOICEOVER_SYNC, but the director reviews
  // scenes in BATCHES — a later batch can't see the lines an earlier batch already
  // used, so the same sentence gets assigned twice (the Kodak cut narrated 6 lines
  // twice). Enforce it in CODE: scan every scene for this idea and alert on exact
  // repeats so they're fixed BEFORE Stage 7B spends ElevenLabs credits.
  const seenVO = {}, dupVO = [];
  const vlAll = visualSheet.getDataRange().getValues();
  for (let i = 1; i < vlAll.length; i++) {
    if (String(vlAll[i][COL_VISUAL.ID - 1]).trim() !== contentId) continue;
    const line = String(vlAll[i][COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1] || "")
      .trim().toLowerCase().replace(/\s+/g, " ");
    if (!line) continue;
    const sn = vlAll[i][COL_VISUAL.SCENE_NUM - 1];
    if (seenVO[line]) dupVO.push("Scene " + sn + " repeats Scene " + seenVO[line]);
    else seenVO[line] = sn;
  }
  if (dupVO.length) {
    SpreadsheetApp.getUi().alert(
      "⚠ DUPLICATE VOICEOVER — FIX BEFORE STAGE 7B",
      dupVO.length + " scene(s) for " + contentId + " carry a voiceover line already used " +
      "by another scene, so the film would narrate it twice:\n\n" + dupVO.join("\n") +
      "\n\nEdit or delete the repeated scene(s) before running Stage 7B / Stage 9C.",
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}