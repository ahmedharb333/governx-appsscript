/* ============================================================================
   Director_Skill.gs — GovernX Content OS
   Stage 4.5 — The Director Pass

   WHAT THIS IS:
   A full creative direction review that runs AFTER Stage 4 (scenes generated)
   and BEFORE Stage 9 (assembly guide built). It reads the complete script +
   scene list together and outputs three things:

   1. VISUAL-NARRATIVE SYNC MAP
      Scene-by-scene verdict on whether each visual earns its voiceover beat.
      Flags weak matches and outputs replacement briefs.

   2. INFOGRAPHIC & DATA VISUALIZATION BRIEF
      Identifies every moment where a stat, comparison, or timeline needs a
      custom graphic rather than B-roll. Outputs precise build instructions
      for each graphic (tool, data, visual form, CapCut integration).

   3. EMOTIONAL PACING & STOP-SCROLLING AUDIT
      Maps the full video arc: tension build, hook impact, checkpoint
      escalation, root cause peak, GRC closing authority. Flags any pacing
      breaks and outputs a revised scene order if needed.

   OUTPUT:
   - Writes a Director Report to the content's Google Drive folder
   - Writes a summary verdict + any infographic scene replacements back
     to the Visual Library (adds new Infographic scene rows where needed)
   - Links the Director Report in Publishing Tracker → Notes column
   - Blocks Stage 9 from running until Director Pass is marked complete

   POSITION IN PIPELINE:
   Stage 4 → Scene generation (existing)
   Stage 4.5 → Director Pass (this file) ← NEW
   Stage 5 → Publishing Row (existing)
   ...
   Stage 9 → Assembly Guide (existing — now reads Director output)
   ============================================================================ */


// ── Director Skill System Prompt ──────────────────────────────────────────────
// This is a SEPARATE system context from SYSTEM_CONTEXT in config.gs.
// The Director is a critic and architect, not a writer.
// It never sees the same prompt as the creator Claude instance.

const DIRECTOR_SYSTEM_CONTEXT = `
You are the Creative Director for GovernX — a YouTube channel that reverse-engineers 
leadership decisions to reveal the hidden role of GRC and BPR in organizational outcomes.

YOUR ROLE IS NOT TO WRITE. IT IS TO JUDGE AND DIRECT.

You review a completed script + scene list as a single unified video experience.
Your job is to decide whether this video will work — as film, not as document.

YOUR THREE RESPONSIBILITIES:

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. VISUAL-NARRATIVE SYNC
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
For each scene, ask: does this visual earn the voiceover beat it covers?

A scene earns its beat when:
- The visual AMPLIFIES the emotional weight of the words (not just decorates them)
- The viewer would feel the stakes visually, even with the sound off
- There is no disconnect between what is said and what is shown

A scene FAILS its beat when:
- Generic B-roll plays over specific, emotionally charged language
- A concrete statistic is spoken but nothing on screen makes it land
- A symbolic moment (root cause reveal, governance failure named) uses
  stock footage that has no metaphorical relationship to the story

When a scene fails, you output a REPLACEMENT BRIEF — specific enough that
a new scene can be generated or sourced immediately from your instructions.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
2. INFORMATION ARCHITECTURE — DATA & INFOGRAPHICS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
You identify every moment in the script where DATA needs to be SHOWN, not just spoken.

The rule is: if a viewer can learn more by seeing it than hearing it, it must be visualized.

MANDATORY VISUALIZATION triggers:
- Any comparison between two entities (Serie A revenue vs Premier League revenue)
- Any trend over time (Nokia market share 1998–2013)
- Any number above $1B, or any percentage that represents dramatic change
- The full checkpoint timeline (already handled by Timeline scene — do not duplicate)
- Any "before vs after" moment in the governance story
- Any risk assessment with multiple risk factors → use risk_matrix
- Any set of 2–4 KPIs or metrics that need to be compared → use kpi_dashboard
- Any compliance score, adoption rate, or governance index → use progress_gauge

For each mandatory visualization, you output an INFOGRAPHIC BRIEF specifying:
- WHAT DATA: exact numbers, labels, time range, units
- VISUAL FORM: choose from the full list below
- TOOL: Remotion (programmatic, high quality) | CapCut (text animation) | Canva (graphic export)
- DURATION: how many seconds it needs on screen
- VOICEOVER SYNC: which sentence(s) it plays under

VISUAL FORM OPTIONS — choose the best fit:
  data_callout      → single large stat with label and context line
  counter_animation → animated number count-up from X to Y
  line_graph        → trend over time with year/value data points
  bar_chart         → comparative bars across categories or years
  split_comparison  → two-column side-by-side with multiple row values
  before_after_card → structured before/after table with verdict
  risk_matrix       → 3×3 likelihood vs impact heat grid (use for risk assessments)
  kpi_dashboard     → 2–4 KPI metric cards with trend arrows (use for performance data)
  progress_gauge    → circular arc gauge(s) for compliance %, scores, rates

TOOL SELECTION RULE:
  → Use Remotion for: risk_matrix, kpi_dashboard, progress_gauge, bar_chart, data_callout,
    counter_animation, line_graph, split_comparison, before_after_card
    (Remotion renders these programmatically — no manual build required)
  → Use CapCut only for: custom motion graphics not covered by the above types
  → Use Canva only for: complex static diagrams that need export as image

When TOOL = Remotion, you MUST output a REMOTION_DATA field using this exact pipe-separated format:
  risk_matrix    : type=RISK_MATRIX | title=... | risks=Label:likelihood(1-3):impact(1-3):highlight(true|false),...
  kpi_dashboard  : type=KPI_DASHBOARD | title=... | layout=2x2 | kpis=Label:Value:up|down|neutral:change:context:highlight,...
  progress_gauge : type=PROGRESS_GAUGE | title=... | variant=single|multi | gauges=Label:value(0-100):unit:context:highlight:threshold,...
  bar_chart      : type=BAR_CHART | title=... | unit=... | points=Label:value:highlight(true|false),...
  data_callout   : type=DATA_CALLOUT | value=... | label=... | context=...
  counter_animation: type=COUNTER_ANIMATION | from=0 | to=... | unit=... | label=...
  line_graph     : type=LINE_GRAPH | label=... | unit=... | points=year:value,...
  split_comparison: type=SPLIT_COMPARISON | left_label=... | left_values=Label:val,... | right_label=... | right_values=Label:val,... | bottom_note=...
  before_after_card: type=BEFORE_AFTER_CARD | before_label=... | before_rows=item→val;;item→val | after_label=... | after_rows=item→val;;item→val | verdict=...

IMPORTANT: In REMOTION_DATA, never use | inside a value. Use ;; to separate items in before_rows/after_rows. Use - instead of : inside label text.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
3. EMOTIONAL PACING & THE STOP-SCROLLING TEST
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
You map the video's emotional arc against this mandatory structure:

SECONDS 0–5:   HOOK — creates urgency. A specific number, date, or outcome
               that makes the viewer stop scrolling. No context needed yet.
SECONDS 5–15:  TENSION BUILD — raise the stakes. What did the world see?
               What does this mean? Why does it matter to this viewer?
SECONDS 15–45: CHECKPOINT DESCENT — walk backwards through the decision chain.
               Each checkpoint should feel heavier than the last.
               The viewer should be building dread, not just collecting facts.
SECONDS 45–60: ROOT CAUSE REVEAL — the moment the system failed.
               This is the emotional peak. The visual must match.
SECONDS 60–75: GRC LESSON — authority, not lecture.
               The viewer should feel they just learned something
               no business school would teach them in this format.
SECONDS 75+:   CTA — earned by the quality of what came before.

For videos longer than 90 seconds, this structure scales proportionally.
The ratios — not the timestamps — are what matter.

For each section you score: STRONG / ADEQUATE / WEAK
A WEAK score requires a specific fix note.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
4. VISUAL DISTINCTION & AUDIENCE MAGNETISM AUDIT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GovernX videos must be INSTANTLY RECOGNIZABLE and IMPOSSIBLE TO IGNORE.
This is not about aesthetics. It is about competitive differentiation.

Every scene list must pass the FOUR-QUALITY TEST:

A. DISTINGUISHABLE
   Would a viewer who has seen one GovernX video recognize this as GovernX within 3 seconds?
   - Checkpoint cards use the GovernX visual identity (#0A0A0A / #FF0000 / Montserrat)
   - Data callout cards have the signature red left-border strip
   - The opening title must establish COMPANY + DISCIPLINE + HOOK in one frame
   - If any scene could belong to a generic corporate explainer — it FAILS this test

B. INFORMATIVE & DATA-RICH
   Does this video earn the trust of a C-Suite executive or board member?
   - Minimum 3 data visualizations per video (bar chart, counter, split comparison)
   - Every claim above $100M must have a visual representation on screen
   - Checkpoint cards must show DATE + EVENT + CONSEQUENCE — not just narrative
   - The root cause reveal must include at least one structural diagram or data card
   - If the video could be described as "talking head + B-roll" — it FAILS this test

C. ANALYTICALLY DISTINCTIVE
   Does this video show something the viewer could not get from reading a Wikipedia article?
   - The reverse-engineering angle must be VISIBLE in at least 2 scenes
   - The GRC/BPR discipline must appear as text on screen at least once
   - The causal chain (decision → consequence → root cause) must be traceable visually
   - If a viewer could watch with the sound off and still understand the governance failure
     then the visual layer is doing its job — otherwise it is FAILING

D. AUDIENCE-ATTRACTING
   Would a CFO, board member, or C-Suite executive stop scrolling for this?
   Standards:
   - The first 3 seconds must contain a specific number, name, or outcome — not a question
   - At least one scene must create a "I never knew that" moment through data
   - The emotional arc must reach a clear peak — there must be a moment of genuine revelation
   - The closing GRC lesson must feel like actionable intelligence, not a moral lecture

FOR EACH OF THE FOUR QUALITIES, score: STRONG / ADEQUATE / WEAK
A WEAK score blocks Stage 9 and requires specific scene additions or replacements.

MANDATORY SCENE ADDITIONS if any quality scores WEAK:
- Missing data visualization → add Infographic scene with specific data brief
- Weak visual identity → add Checkpoint card with GovernX styling
- Weak reverse-engineering visibility → add Text Impact scene naming the root cause
- Weak hook power → flag the opening scene for replacement with a stat-forward brief

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
YOUR TONE AS DIRECTOR:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
You are direct, specific, and constructive.
You do not soften failures — you name them precisely.
You do not praise what works without explaining WHY it works.
When something needs to change, you say EXACTLY what to change and how.
Your output is a production instruction, not a review essay.

GOVERNX VISUAL IDENTITY (always enforce):
- Background: #0A0A0A near-black
- Accent: #FF0000 vivid red
- Primary text: #FFFFFF
- Secondary text: #CCCCCC
- Fonts: Montserrat (English) / Cairo (Arabic)
- No gradients, no soft colors, no decorative elements
- Every visual either advances the story or it does not exist
`;


// ══════════════════════════════════════════════════════════════════════════════
// STAGE 4.5 — Run the Director Pass
// Called from GovernX menu after Stage 4 completes
// ══════════════════════════════════════════════════════════════════════════════
function runDirectorPass() {
  const idea = getActiveIdeaRow();
  if (!idea) return;

  // Director requires both script (Stage 3) and scenes (Stage 4)
  if (!checkPreviousStage(idea.id, SHEET.SCRIPT, "Stage 3 — Script")) return;
  if (!checkPreviousStage(idea.id, SHEET.VISUAL, "Stage 4 — Scenes")) return;

  const ui     = SpreadsheetApp.getUi();
  const master = getMasterContent(idea.id);
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  // ── Pull script ─────────────────────────────────────────────────────────────
  const scriptSheet = ss.getSheetByName(SHEET.SCRIPT);
  const scriptData  = scriptSheet.getDataRange().getValues();
  let   script      = null;

  for (let i = 1; i < scriptData.length; i++) {
    if (scriptData[i][COL_SCRIPT.ID - 1].toString().trim() === idea.id) {
      script = {
        hook      : scriptData[i][COL_SCRIPT.HOOK             - 1],
        narrative : scriptData[i][COL_SCRIPT.NARRATIVE_FLOW   - 1],
        voiceover : scriptData[i][COL_SCRIPT.VOICEOVER_SCRIPT - 1],
        closing   : scriptData[i][COL_SCRIPT.GRC_BPR_CLOSING  - 1],
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

  // ── Pull scenes ──────────────────────────────────────────────────────────────
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    scenes.push({
      row             : i + 1,
      num             : visualData[i][COL_VISUAL.SCENE_NUM        - 1],
      type            : visualData[i][COL_VISUAL.SCENE_TYPE       - 1],
      description     : visualData[i][COL_VISUAL.DESCRIPTION      - 1],
      source          : visualData[i][COL_VISUAL.SOURCE           - 1],
      link            : visualData[i][COL_VISUAL.LINK             - 1],
      timestamp       : visualData[i][COL_VISUAL.TIMESTAMP        - 1],
      license         : visualData[i][COL_VISUAL.LICENSE          - 1],
      assemblyNotes   : visualData[i][COL_VISUAL.ASSEMBLY_NOTES   - 1],
      checkpointDate  : visualData[i][COL_VISUAL.CHECKPOINT_DATE  - 1],
      checkpointEvent : visualData[i][COL_VISUAL.CHECKPOINT_EVENT - 1],
      checkpointAngle : visualData[i][COL_VISUAL.CHECKPOINT_ANGLE - 1]
    });
  }

  if (scenes.length === 0) {
    ui.alert("No scenes found for: " + idea.id + "\nRun Stage 4 first.");
    return;
  }

  // ── Pull structured research data (Risk/KPI/Gauge rows from Stage 2) ────────
  const researchData = getStructuredResearchData_(ss, idea.id);

  // ── Stage 2 sufficiency check ─────────────────────────────────────────────
  const totalStructured = researchData.risks.length + researchData.kpis.length + researchData.gauges.length;
  if (totalStructured === 0) {
    const proceed = ui.alert(
      "⚠️ No Structured Visualization Data",
      "The Research Database has no Risk, KPI, or Gauge rows for: " + idea.id + "\n\n" +
      "This means the Director will generate estimated numbers for Risk Matrix,\n" +
      "KPI Dashboard, and Progress Gauge compositions — not sourced real data.\n\n" +
      "OPTIONS:\n" +
      "• Re-run Stage 2 to collect structured data, then re-run Director Pass\n" +
      "• Continue anyway — Director will use script context to estimate values\n\n" +
      "Continue without structured data?",
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }

  ui.alert(
    "Stage 4.5 — Director Pass",
    "The Director is reviewing the full script + scene list for: " + idea.company +
    "\n\nThis evaluates:\n" +
    "• Visual-narrative sync (scene by scene)\n" +
    "• Data visualization & infographic opportunities\n" +
    "• Emotional pacing & stop-scrolling audit\n\n" +
    "Research data loaded: " +
    researchData.risks.length + " risks, " +
    researchData.kpis.length + " KPIs, " +
    researchData.gauges.length + " gauges" +
    (totalStructured === 0 ? " ⚠️ (none — Director will estimate)" : " ✅") +
    "\n\nThis may take 30–45 seconds.",
    ui.ButtonSet.OK
  );

  try {
    // ── Build Director prompt ───────────────────────────────────────────────
    const prompt = buildDirectorPrompt(idea, master, script, scenes, researchData);

    // ── Call Claude with Director system context ────────────────────────────
    // NOTE: uses callClaudeAsDirector — separate from callClaude()
    // The Director NEVER uses SYSTEM_CONTEXT from config.gs
    const raw = callClaudeAsDirector(prompt);

    // ── Parse the Director output ───────────────────────────────────────────
    const directorReport = parseDirectorOutput(raw);

    // ── Write infographic scenes back to Visual Library ─────────────────────
    const infographicsAdded = writeInfographicScenes(idea, directorReport, visualSheet, scenes);

    // ── Force Director verdicts into Visual Library ─────────────────────────
    // FAIL/WEAK scenes get their description and assembly notes updated
    // so Stage 8 picks up the Director's replacement brief automatically
    const verdictsApplied = applyDirectorVerdicts(idea, directorReport, visualSheet, scenes);

    // ── Apply mandatory additions from WEAK distinction scores ───────────────
    // Adds new scenes to Visual Library for any WEAK quality dimension
    applyDistinctionAdditions(idea, directorReport, visualSheet);

    // ── Export Director Report to Google Drive ──────────────────────────────
    const reportUrl = exportDirectorReport(idea, master, raw, directorReport, infographicsAdded);

    // ── Write report link to Publishing Tracker Notes ───────────────────────
    writeDirectorLinkToPublishing(idea.id, idea.company, master, reportUrl);

    // ── Summary alert ────────────────────────────────────────────────────────
    const verdict = directorReport.overallVerdict || "REVIEW REQUIRED";
    const syncFails    = directorReport.syncFailures    || 0;
    const infographics = infographicsAdded               || 0;
    const pacingWeak   = directorReport.pacingWeakCount  || 0;

    let summary = "Director Pass Complete for: " + idea.id + "\n\n";
    summary += "OVERALL VERDICT: " + verdict + "\n\n";
    summary += "📊 Infographic scenes added : " + infographics + "\n";
    summary += "🔁 Visual sync failures    : " + syncFails + "\n";
    summary += "✏️  Scene descriptions updated: " + verdictsApplied + " (FAIL/WEAK replaced in Visual Library)\n";
    summary += "⚡ Pacing sections weak    : " + pacingWeak + " (see report)\n";

    // Distinction scores
    if (directorReport.distinctionScores && directorReport.distinctionScores.length > 0) {
      summary += "\n🎯 VISUAL QUALITY SCORES:\n";
      directorReport.distinctionScores.forEach(d => {
        const icon = d.score === "STRONG" ? "✅" : d.score === "ADEQUATE" ? "⚡" : "❌";
        summary += icon + " " + d.quality + ": " + d.score + "\n";
      });
      if (directorReport.distinctionWeakCount > 0) {
        summary += "\n⚠️ " + directorReport.distinctionWeakCount + " quality dimension(s) WEAK — mandatory scenes added to Visual Library.\n";
      }
    }

    summary += "\n";
    summary += "Director Report saved to Drive:\n" + reportUrl + "\n\n";

    if (verdict === "APPROVED") {
      summary += "✅ Video approved for production.\nProceed to Stage 5.";
    } else if (verdict === "APPROVED WITH NOTES") {
      summary += "⚠️ Approved with notes. Review Director Report before Stage 9 (Assembly Guide).";
    } else {
      summary += "⛔ Revisions required before proceeding.\nReview Director Report, fix flagged scenes, then re-run Stage 4.5.";
    }

    updatePipelineStatus_(idea.id, "S45", verdict === "APPROVED" ? "✅" : verdict === "APPROVED WITH NOTES" ? "⚠️" : "❌");
    ui.alert(
      verdict === "APPROVED" ? "✅ Stage 4.5 — Director Approved" :
      verdict === "APPROVED WITH NOTES" ? "⚠️ Stage 4.5 — Approved With Notes" :
      "⛔ Stage 4.5 — Revisions Required",
      summary,
      ui.ButtonSet.OK
    );

  } catch (err) {
    updatePipelineStatus_(idea.id, "S45", "❌");
    logError("Stage 4.5 — Director Pass", idea.id, "Director API Error", err.message);
    ui.alert("❌ Stage 4.5 Failed", err.message + "\nSee Error Log tab.", ui.ButtonSet.OK);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// DIRECTOR PROMPT BUILDER
// Sends complete script + scenes to the Director for evaluation
// ══════════════════════════════════════════════════════════════════════════════
function buildDirectorPrompt(idea, master, script, scenes, researchData) {

  // ── Build scene manifest ─────────────────────────────────────────────────
  const sceneManifest = scenes.map(s => {
    let line = `SCENE ${s.num} | ${s.timestamp} | ${s.type.toUpperCase()} | Source: ${s.source}`;
    if (s.type === "Checkpoint") {
      line += `\n  Date: ${s.checkpointDate} | Event: ${s.checkpointEvent} | Angle: ${s.checkpointAngle}`;
    } else {
      line += `\n  Visual: ${s.description}`;
    }
    return line;
  }).join("\n\n");

  // ── Trim voiceover for context (first 1200 chars covers the opening arc) ──
  const voiceoverPreview = script.voiceover
    ? script.voiceover.substring(0, 1500) + (script.voiceover.length > 1500 ? "\n[...continues]" : "")
    : "(No voiceover script found)";

  return `
You are directing a GovernX video. Review the complete script and scene list below.

═══════════════════════════════════════════════════
VIDEO BRIEF
═══════════════════════════════════════════════════
Content ID    : ${idea.id}
Title         : ${master ? master.title : idea.company}
Company/Topic : ${idea.company}
Discipline    : ${master ? master.discipline : "GRC"}
Language      : ${idea.language}
Format        : ${script.format}
Checkpoints   : ${master ? master.checkpoints : "Not available"}
Core Insight  : ${master ? master.coreInsight : "Not available"}

═══════════════════════════════════════════════════
FULL VOICEOVER SCRIPT
═══════════════════════════════════════════════════
HOOK:
${script.hook}

NARRATIVE FLOW:
${script.narrative}

VOICEOVER SCRIPT:
${voiceoverPreview}

GRC/BPR CLOSING:
${script.closing}

═══════════════════════════════════════════════════
SCENE LIST (${scenes.length} scenes total)
═══════════════════════════════════════════════════
${sceneManifest}

${buildResearchDataSection_(researchData)}
═══════════════════════════════════════════════════
YOUR DIRECTOR REVIEW — RETURN IN EXACTLY THIS FORMAT
═══════════════════════════════════════════════════

DIRECTOR_REPORT_START

OVERALL_VERDICT: [APPROVED | APPROVED WITH NOTES | REVISIONS REQUIRED]
VERDICT_REASON: [1-2 sentences — what drives this verdict]

─────────────────────────────────────────
SECTION A — VISUAL-NARRATIVE SYNC
─────────────────────────────────────────
For EVERY scene, return one SYNC_SCENE block.
Scenes that pass get a one-line confirmation.
Scenes that fail get a full REPLACEMENT_BRIEF.

SYNC_SCENE_START
SCENE_NUM: [number]
SCENE_TYPE: [type]
SYNC_VERDICT: [PASS | FAIL | WEAK]
SYNC_REASON: [Why this verdict — be specific about the voiceover beat it covers]
REPLACEMENT_BRIEF: [ONLY if FAIL or WEAK — exact visual replacement instruction.
  Include: scene type, visual concept, source recommendation (Pexels/KlingAI/Veo/CapCut),
  specific prompt or search query, and which voiceover sentence(s) it must cover]
REMOTION_DATA: [For Remotion scenes ONLY — exact pipe-separated key=value string using the formats above.
  For non-Remotion scenes (B-roll, Checkpoint, Text) write: N/A]
REMOTION_STYLE: [Template name and variant e.g. "InfographicScene | bar_chart | weight=bold" — write N/A if not Remotion]
VOICEOVER_SYNC: [Exact sentence(s) from the voiceover script that play during this scene. Copy verbatim.]
SCENE_SCORE: [A | B | NEEDS_DATA]
SYNC_SCENE_END

[repeat for every scene]

─────────────────────────────────────────
SECTION B — INFOGRAPHIC & DATA BRIEFS
─────────────────────────────────────────
List every moment that requires a custom data visualization.
If none are needed, write: NO_INFOGRAPHICS_NEEDED

INFOGRAPHIC_START
INFOGRAPHIC_NUM: [sequential number]
TRIGGER_MOMENT: [Quote the exact sentence or phrase from the voiceover that triggers this]
VOICEOVER_SYNC: [Which sentence(s) this graphic plays under — copy exact text]
DATA_REQUIRED: [Exact numbers, labels, time range, units needed]
VISUAL_FORM: [data_callout | counter_animation | line_graph | bar_chart | split_comparison | before_after_card | risk_matrix | kpi_dashboard | progress_gauge]
TOOL: [Remotion | CapCut | Canva]
REMOTION_DATA: [ONLY if TOOL=Remotion — exact pipe-separated key=value string. Example: type=BAR_CHART | title=Pension Deficit Growth | unit=£M | points=2000:190:false,2010:345:false,2016:571:true]
DURATION_SECONDS: [number]
INSERT_AFTER_SCENE: [scene number this graphic should follow OR replace]
ACTION: [INSERT_NEW_SCENE | REPLACE_SCENE_[N]]
CAPCUT_BUILD_INSTRUCTIONS: [ONLY if TOOL=CapCut or Canva — step-by-step build instructions. Write N/A if TOOL=Remotion]
INFOGRAPHIC_END

[repeat for each infographic needed]

─────────────────────────────────────────
SECTION C — EMOTIONAL PACING AUDIT
─────────────────────────────────────────
PACING_SECTION_START
SECTION_NAME: HOOK (0–5s)
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [What works or what fails — be specific]
FIX_REQUIRED: [Only if WEAK — exact instruction to fix]
PACING_SECTION_END

PACING_SECTION_START
SECTION_NAME: TENSION_BUILD (5–15s)
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [assessment]
FIX_REQUIRED: [fix if WEAK]
PACING_SECTION_END

PACING_SECTION_START
SECTION_NAME: CHECKPOINT_DESCENT
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [assessment]
FIX_REQUIRED: [fix if WEAK]
PACING_SECTION_END

PACING_SECTION_START
SECTION_NAME: ROOT_CAUSE_REVEAL
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [assessment]
FIX_REQUIRED: [fix if WEAK]
PACING_SECTION_END

PACING_SECTION_START
SECTION_NAME: GRC_CLOSING
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [assessment]
FIX_REQUIRED: [fix if WEAK]
PACING_SECTION_END

─────────────────────────────────────────
SECTION D — VISUAL DISTINCTION & AUDIENCE MAGNETISM
─────────────────────────────────────────
Score each of the four qualities. WEAK on any requires mandatory scene additions.

DISTINCTION_START
QUALITY: DISTINGUISHABLE
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [Is this instantly recognizable as GovernX? What confirms or breaks identity?]
MANDATORY_ADDITION: [Only if WEAK — exact scene to add with full brief]
DISTINCTION_END

DISTINCTION_START
QUALITY: DATA_RICH
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [How many data visualizations? Are all major claims visualized? What is missing?]
MANDATORY_ADDITION: [Only if WEAK — exact infographic scene brief with data and visual form]
DISTINCTION_END

DISTINCTION_START
QUALITY: ANALYTICALLY_DISTINCTIVE
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [Is the reverse-engineering angle visible on screen? Is GRC discipline named visually?]
MANDATORY_ADDITION: [Only if WEAK — exact text impact card or scene brief]
DISTINCTION_END

DISTINCTION_START
QUALITY: AUDIENCE_ATTRACTING
SCORE: [STRONG | ADEQUATE | WEAK]
ASSESSMENT: [Would a CFO or board member stop scrolling? Is there a genuine revelation moment?]
MANDATORY_ADDITION: [Only if WEAK — exact fix instruction for hook or revelation scene]
DISTINCTION_END

─────────────────────────────────────────
SECTION E — DIRECTOR'S FINAL VERDICT
─────────────────────────────────────────
FINAL_NOTES: [
  3–5 sentences.
  What is the single strongest element of this video?
  What is the one change with the highest impact on audience retention?
  Does this video meet the GovernX standard: distinguishable, informative,
  analytically distinctive, and audience-attracting?
  Is this video ready to represent the GovernX brand to a C-Suite audience?
]

DIRECTOR_REPORT_END
`;
}


// ══════════════════════════════════════════════════════════════════════════════
// CALL CLAUDE AS DIRECTOR
// Uses DIRECTOR_SYSTEM_CONTEXT — completely separate from SYSTEM_CONTEXT
// Higher token budget — Director needs space to analyze every scene in detail
// ══════════════════════════════════════════════════════════════════════════════
function callClaudeAsDirector(finalPrompt) {

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
        max_tokens : 12000,        // Director needs more tokens — reviewing every scene
        system     : DIRECTOR_SYSTEM_CONTEXT,   // ← Director context, NOT SYSTEM_CONTEXT
        messages   : [
          { role: "user", content: finalPrompt }
        ]
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

      if (code === 529) {
        lastError = "API overloaded (529)";
        if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
        continue;
      }

      if (code === 429) {
        lastError = "Rate limited (429)";
        if (attempt < MAX_TRIES) Utilities.sleep(RATELIMIT_MS);
        continue;
      }

      throw new Error("Director API error " + code + ": " + body);

    } catch (err) {
      lastError = err.message;
      if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
    }
  }

  throw new Error(
    "Director API failed after " + MAX_TRIES +
    " attempts. Last error: " + lastError
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// RESEARCH DATA HELPERS — Pull structured Risk/KPI/Gauge rows for Director prompt
// ══════════════════════════════════════════════════════════════════════════════

function getStructuredResearchData_(ss, contentId) {
  const result = { risks: [], kpis: [], gauges: [] };
  const sheet  = ss.getSheetByName(SHEET.RESEARCH);
  if (!sheet) return result;

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_RESEARCH.ID - 1].toString().trim() !== contentId) continue;
    const sourceType = data[i][COL_RESEARCH.SOURCE_TYPE - 1].toString().trim();
    const note       = data[i][COL_RESEARCH.NOTE        - 1].toString().trim();
    if (!note) continue;

    // Parse pipe-separated key:value pairs from NOTE column
    const fields = {};
    note.split("|").forEach(pair => {
      const idx = pair.indexOf(":");
      if (idx < 0) return;
      fields[pair.substring(0, idx).trim()] = pair.substring(idx + 1).trim();
    });

    if (sourceType === "Risk")  result.risks.push(fields);
    if (sourceType === "KPI")   result.kpis.push(fields);
    if (sourceType === "Gauge") result.gauges.push(fields);
  }
  return result;
}

function buildResearchDataSection_(researchData) {
  if (!researchData) return "";
  const { risks, kpis, gauges } = researchData;
  if (risks.length + kpis.length + gauges.length === 0) return "";

  let section = `═══════════════════════════════════════════════════
STRUCTURED RESEARCH DATA — Use these real numbers in REMOTION_DATA
═══════════════════════════════════════════════════
These are sourced, verified data points from Stage 2 research.
You MUST prioritise these when generating REMOTION_DATA for Risk Matrix,
KPI Dashboard, and Progress Gauge visualizations. Do not invent values
when real ones are available here.

`;

  if (risks.length > 0) {
    section += "RISK FACTORS:\n";
    risks.forEach((r, i) => {
      section += `  ${i + 1}. ${r["RISK_LABEL"] || "?"} — Likelihood:${r["RISK_LIKELIHOOD"] || "?"} / Impact:${r["RISK_IMPACT"] || "?"} | Highlight:${r["RISK_HIGHLIGHT"] || "false"}\n`;
      if (r["RISK_DESCRIPTION"]) section += `     ${r["RISK_DESCRIPTION"]}\n`;
    });
    section += "\n";
  }

  if (kpis.length > 0) {
    section += "KEY PERFORMANCE METRICS:\n";
    kpis.forEach((k, i) => {
      section += `  ${i + 1}. ${k["KPI_LABEL"] || "?"}: ${k["KPI_VALUE"] || "?"} (${k["KPI_TREND"] || "?"} ${k["KPI_CHANGE"] || ""}) — ${k["KPI_CONTEXT"] || ""} | Highlight:${k["KPI_HIGHLIGHT"] || "false"}\n`;
    });
    section += "\n";
  }

  if (gauges.length > 0) {
    section += "COMPLIANCE / COVERAGE GAUGES:\n";
    gauges.forEach((g, i) => {
      section += `  ${i + 1}. ${g["GAUGE_LABEL"] || "?"}: ${g["GAUGE_VALUE"] || "?"}${g["GAUGE_UNIT"] || "%"} — ${g["GAUGE_CONTEXT"] || ""} | Threshold:${g["GAUGE_THRESHOLD"] || "N/A"} | Highlight:${g["GAUGE_HIGHLIGHT"] || "false"}\n`;
    });
    section += "\n";
  }

  return section;
}

// ══════════════════════════════════════════════════════════════════════════════
// PARSE DIRECTOR OUTPUT
// Extracts structured data from the raw Director report text
// ══════════════════════════════════════════════════════════════════════════════
function parseDirectorOutput(raw) {

  const result = {
    overallVerdict      : "REVIEW REQUIRED",
    verdictReason       : "",
    syncFailures        : 0,
    syncScenes          : [],
    infographics        : [],
    pacingSections      : [],
    pacingWeakCount     : 0,
    distinctionScores   : [],   // NEW — Visual Distinction audit
    distinctionWeakCount: 0,    // NEW — count of WEAK distinction scores
    finalNotes          : ""
  };

  // ── Overall verdict ────────────────────────────────────────────────────────
  const verdictMatch = raw.match(/OVERALL_VERDICT:\s*(.+)/);
  if (verdictMatch) result.overallVerdict = verdictMatch[1].trim();

  const reasonMatch = raw.match(/VERDICT_REASON:\s*(.+)/);
  if (reasonMatch) result.verdictReason = reasonMatch[1].trim();

  // ── Sync scenes ────────────────────────────────────────────────────────────
  const syncBlocks = raw.match(/SYNC_SCENE_START([\s\S]*?)SYNC_SCENE_END/g) || [];
  syncBlocks.forEach(block => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };
    const getBlock = (field) => {
      const m = block.match(new RegExp(field + ":\\s*([\\s\\S]*?)(?=\\n[A-Z_]+:|SYNC_SCENE_END)"));
      return m ? m[1].trim() : "";
    };

    const verdict = get("SYNC_VERDICT");
    const remotionData = get("REMOTION_DATA");
    const scene = {
      sceneNum          : get("SCENE_NUM"),
      sceneType         : get("SCENE_TYPE"),
      syncVerdict       : verdict,
      syncReason        : get("SYNC_REASON"),
      replacementBrief  : getBlock("REPLACEMENT_BRIEF"),
      remotionData      : remotionData === "N/A" ? "" : remotionData,
      remotionStyle     : (() => { const v = get("REMOTION_STYLE"); return v === "N/A" ? "" : v; })(),
      voiceoverSync     : getBlock("VOICEOVER_SYNC"),
      sceneScore        : get("SCENE_SCORE")
    };

    result.syncScenes.push(scene);
    if (verdict === "FAIL" || verdict === "WEAK") result.syncFailures++;
  });

  // ── Infographics ───────────────────────────────────────────────────────────
  const infoBlocks = raw.match(/INFOGRAPHIC_START([\s\S]*?)INFOGRAPHIC_END/g) || [];
  infoBlocks.forEach(block => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };
    const getBlock = (field) => {
      const m = block.match(new RegExp(field + ":\\s*\\[?([\\s\\S]*?)(?=\\n[A-Z_]+:|INFOGRAPHIC_END)"));
      return m ? m[1].replace(/^\[|\]$/g, "").trim() : "";
    };

    result.infographics.push({
      num                   : get("INFOGRAPHIC_NUM"),
      triggerMoment         : get("TRIGGER_MOMENT"),
      voiceoverSync         : get("VOICEOVER_SYNC"),
      dataRequired          : get("DATA_REQUIRED"),
      visualForm            : get("VISUAL_FORM"),
      tool                  : get("TOOL"),
      remotionData          : get("REMOTION_DATA"),
      durationSeconds       : get("DURATION_SECONDS"),
      insertAfterScene      : get("INSERT_AFTER_SCENE"),
      action                : get("ACTION"),
      capCutInstructions    : getBlock("CAPCUT_BUILD_INSTRUCTIONS")
    });
  });

  // ── Pacing sections ────────────────────────────────────────────────────────
  const pacingBlocks = raw.match(/PACING_SECTION_START([\s\S]*?)PACING_SECTION_END/g) || [];
  pacingBlocks.forEach(block => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };

    const score = get("SCORE");
    result.pacingSections.push({
      name       : get("SECTION_NAME"),
      score      : score,
      assessment : get("ASSESSMENT"),
      fix        : get("FIX_REQUIRED")
    });
    if (score === "WEAK") result.pacingWeakCount++;
  });

  // ── Distinction scores (Section D) ────────────────────────────────────────
  const distinctionBlocks = raw.match(/DISTINCTION_START([\s\S]*?)DISTINCTION_END/g) || [];
  distinctionBlocks.forEach(block => {
    const get = (field) => {
      const m = block.match(new RegExp(field + ":\\s*(.+)"));
      return m ? m[1].trim() : "";
    };
    const getBlock = (field) => {
      const m = block.match(new RegExp(field + ":\\s*\\[?([\\s\\S]*?)(?=\\n[A-Z_]+:|DISTINCTION_END)"));
      return m ? m[1].replace(/^\[|\]$/g, "").trim() : "";
    };

    const score = get("SCORE");
    result.distinctionScores.push({
      quality           : get("QUALITY"),
      score             : score,
      assessment        : get("ASSESSMENT"),
      mandatoryAddition : getBlock("MANDATORY_ADDITION")
    });
    if (score === "WEAK") result.distinctionWeakCount++;
  });

  // ── Final notes ────────────────────────────────────────────────────────────
  const notesMatch = raw.match(/FINAL_NOTES:\s*\[?([\s\S]*?)(?=\]?\s*DIRECTOR_REPORT_END)/);
  if (notesMatch) result.finalNotes = notesMatch[1].trim();

  return result;
}


// ══════════════════════════════════════════════════════════════════════════════
// APPLY DISTINCTION ADDITIONS TO VISUAL LIBRARY
// For every WEAK quality dimension, adds the Director's mandatory scene
// to the Visual Library so Stage 8 builds it automatically
// ══════════════════════════════════════════════════════════════════════════════
function applyDistinctionAdditions(idea, directorReport, visualSheet) {

  if (!directorReport.distinctionScores || directorReport.distinctionWeakCount === 0) return;

  const weakScores = directorReport.distinctionScores.filter(
    d => d.score === "WEAK" && d.mandatoryAddition && d.mandatoryAddition.trim() !== ""
  );

  if (weakScores.length === 0) return;

  // Find the last row for this content ID
  const data       = visualSheet.getDataRange().getValues();
  let   lastRow    = 1;
  let   lastSceneNum = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_VISUAL.ID - 1].toString().trim() === idea.id) {
      lastRow = i + 1;
      const num = parseInt(data[i][COL_VISUAL.SCENE_NUM - 1].toString().replace(/\D/g, "")) || 0;
      if (num > lastSceneNum) lastSceneNum = num;
    }
  }

  weakScores.forEach((d, idx) => {
    const newRow    = lastRow + 1 + idx;
    const sceneNum  = "D" + (idx + 1); // D1, D2 = Distinction additions

    // Determine scene type from quality
    const typeMap = {
      "DISTINGUISHABLE"          : "Checkpoint",
      "DATA_RICH"                : "KPI Dashboard",
      "ANALYTICALLY_DISTINCTIVE" : "Text",
      "AUDIENCE_ATTRACTING"      : "Text"
    };
    const sceneType = typeMap[d.quality] || "Text";

    visualSheet.getRange(newRow, COL_VISUAL.ID         ).setValue(idea.id);
    visualSheet.getRange(newRow, COL_VISUAL.SCENE_NUM  ).setValue(sceneNum);
    // Clear validation on SCENE_TYPE cell before writing — prevents dropdown rejection
    visualSheet.getRange(newRow, COL_VISUAL.SCENE_TYPE ).clearDataValidations();
    visualSheet.getRange(newRow, COL_VISUAL.SCENE_TYPE ).setValue(sceneType);
    visualSheet.getRange(newRow, COL_VISUAL.DESCRIPTION).setValue(
      "⚠️ DIRECTOR ADDITION — " + d.quality + " WEAK\n" + d.mandatoryAddition
    );
    // Clear validation on STATUS cell too — same dropdown issue
    visualSheet.getRange(newRow, COL_VISUAL.STATUS     ).clearDataValidations();
    visualSheet.getRange(newRow, COL_VISUAL.STATUS     ).setValue("Needed");
    visualSheet.getRange(newRow, COL_VISUAL.ASSEMBLY_NOTES).setValue(
      "Added by Director Pass — Quality dimension: " + d.quality + "\n" + d.assessment
    );

    // Amber background — mandatory Director addition
    visualSheet.getRange(newRow, 1, 1, 18).setBackground("#FFF3CD");

    Logger.log("Distinction addition: " + d.quality + " → Scene " + sceneNum + " added");
  });

  SpreadsheetApp.flush();
  Logger.log("applyDistinctionAdditions: " + weakScores.length + " scenes added");
}


// ══════════════════════════════════════════════════════════════════════════════
// APPLY DIRECTOR VERDICTS TO VISUAL LIBRARY
// For every scene flagged FAIL or WEAK:
//   - Updates DESCRIPTION with the Director's replacement brief
//   - Updates ASSEMBLY_NOTES with Director reasoning
//   - Sets STATUS back to "Needed" so Stage 8 re-processes it
//   - Highlights row in amber (WEAK) or red (FAIL)
// This means Stage 8 automatically uses the Director's corrected brief
// ══════════════════════════════════════════════════════════════════════════════
function applyDirectorVerdicts(idea, directorReport, visualSheet, scenes) {

  if (!directorReport.syncScenes || directorReport.syncScenes.length === 0) return 0;

  // Build a lookup: sceneNum → scene row index in sheet
  const sceneRowMap = {};
  const data = visualSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    const num = data[i][COL_VISUAL.SCENE_NUM - 1].toString().trim();
    sceneRowMap[num] = i + 1; // 1-based sheet row
  }

  let updated = 0;

  // Process ALL scenes — PASS scenes get cols 19-22 written, FAIL/WEAK also get description updated
  directorReport.syncScenes.forEach(scene => {
    const sheetRow = sceneRowMap[scene.sceneNum.toString()];
    if (!sheetRow) {
      Logger.log("Director verdict: scene " + scene.sceneNum + " not found in Visual Library");
      return;
    }

    const isFail = scene.syncVerdict === "FAIL" || scene.syncVerdict === "WEAK";

    // ── Write cols 19-22 (S-V) for ALL scenes ────────────────────────────────
    if (scene.remotionData && scene.remotionData.trim()) {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.REMOTION_DATA)
        .setValue(scene.remotionData.trim());
    }
    if (scene.remotionStyle && scene.remotionStyle.trim()) {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.REMOTION_STYLE)
        .setValue(scene.remotionStyle.trim());
    }
    if (scene.voiceoverSync && scene.voiceoverSync.trim()) {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.VOICEOVER_SYNC)
        .setValue(scene.voiceoverSync.trim());
    }
    if (scene.sceneScore && scene.sceneScore.trim()) {
      visualSheet.getRange(sheetRow, COL_VISUAL_EXTENDED.SCENE_SCORE)
        .setValue(scene.sceneScore.trim());
    }

    // ── FAIL / WEAK only: update description + assembly notes + status + color ─
    if (isFail) {
      if (scene.replacementBrief && scene.replacementBrief.trim() !== "") {
        visualSheet.getRange(sheetRow, COL_VISUAL.DESCRIPTION)
          .setValue(scene.replacementBrief.trim());
      }

      const directorNote = "⚠️ DIRECTOR " + scene.syncVerdict + ": " + scene.syncReason +
        (scene.replacementBrief ? "\nREPLACEMENT: " + scene.replacementBrief : "");
      const currentNotes = visualSheet.getRange(sheetRow, COL_VISUAL.ASSEMBLY_NOTES).getValue();
      visualSheet.getRange(sheetRow, COL_VISUAL.ASSEMBLY_NOTES)
        .setValue(directorNote + (currentNotes ? "\n---\n" + currentNotes : ""));

      const currentStatus = visualSheet.getRange(sheetRow, COL_VISUAL.STATUS).getValue();
      if (currentStatus !== "Done" && currentStatus !== "Skip") {
        visualSheet.getRange(sheetRow, COL_VISUAL.STATUS).clearDataValidations();
        visualSheet.getRange(sheetRow, COL_VISUAL.STATUS).setValue("Needed");
      }

      const color = scene.syncVerdict === "FAIL" ? "#FFD7D7" : "#FFF3CD";
      visualSheet.getRange(sheetRow, 1, 1, 18).setBackground(color);
    }

    updated++;
    Logger.log("Director: Scene " + scene.sceneNum + " (" + scene.syncVerdict +
      ") → S-V written" + (isFail ? ", description + status updated" : ""));
  });

  SpreadsheetApp.flush();
  Logger.log("Director: " + updated + " scenes processed (cols 19-22 written for all)");
  return updated;
}


// ══════════════════════════════════════════════════════════════════════════════
// WRITE INFOGRAPHIC SCENES TO VISUAL LIBRARY
// Inserts new Infographic scene rows after the specified scenes
// Adds them with status "Needed" so they flow into Stage 8A processing
// ══════════════════════════════════════════════════════════════════════════════
function writeInfographicScenes(idea, directorReport, visualSheet, existingScenes) {

  if (!directorReport.infographics || directorReport.infographics.length === 0) {
    return 0;
  }

  let added = 0;

  // Process in reverse order to preserve row positions when inserting
  const sortedInfos = [...directorReport.infographics]
    .filter(info => info.action && info.action.includes("INSERT"))
    .sort((a, b) => {
      const numA = parseInt(a.insertAfterScene) || 0;
      const numB = parseInt(b.insertAfterScene) || 0;
      return numB - numA; // reverse order for insertion
    });

  sortedInfos.forEach(info => {
    try {
      // Find the row of the scene this goes after
      const targetSceneNum = parseInt(info.insertAfterScene) || 0;
      let   insertAfterRow = -1;

      const vData = visualSheet.getDataRange().getValues();
      for (let i = 1; i < vData.length; i++) {
        if (vData[i][COL_VISUAL.ID      - 1].toString().trim() === idea.id &&
            parseInt(vData[i][COL_VISUAL.SCENE_NUM - 1]) === targetSceneNum) {
          insertAfterRow = i + 1;
          break;
        }
      }

      // If we found the target row, insert after it
      // Otherwise append to the end of this content ID's scenes
      let writeRow;
      if (insertAfterRow > 0) {
        visualSheet.insertRowAfter(insertAfterRow);
        writeRow = insertAfterRow + 1;
      } else {
        writeRow = visualSheet.getLastRow() + 1;
      }

      // Clear validations on all columns we'll write to
      const colsToWrite = [
        COL_VISUAL.ID, COL_VISUAL.SCENE_NUM, COL_VISUAL.SCENE_TYPE,
        COL_VISUAL.DESCRIPTION, COL_VISUAL.SOURCE, COL_VISUAL.LINK,
        COL_VISUAL.LICENSE, COL_VISUAL.STATUS, COL_VISUAL.BUILT_WHERE,
        COL_VISUAL.ASSEMBLY_NOTES, COL_VISUAL.VEO_PROMPT
      ];
      colsToWrite.forEach(col => {
        visualSheet.getRange(writeRow, col).clearDataValidations();
      });

      // ── Determine scene type from visual form ──────────────────────────────
      const remotionTypeMap = {
        "risk_matrix"    : "Risk Matrix",
        "kpi_dashboard"  : "KPI Dashboard",
        "progress_gauge" : "Gauge"
      };
      const isRemotionScene = info.tool && info.tool.toLowerCase() === "remotion";
      const sceneType = isRemotionScene && remotionTypeMap[info.visualForm.toLowerCase()]
        ? remotionTypeMap[info.visualForm.toLowerCase()]
        : "Infographic";

      // ── Scene description — what the infographic shows ─────────────────────
      const description =
        "[" + sceneType.toUpperCase() + "] " + info.visualForm.toUpperCase() +
        " — " + info.dataRequired +
        " | Voiceover sync: " + info.voiceoverSync.substring(0, 100);

      // ── Assembly notes — build instructions ────────────────────────────────
      const assemblyNotes = isRemotionScene
        ? "REMOTION RENDER — Stage 8D will build this automatically.\n" +
          "Visual form: " + info.visualForm + "\n" +
          "Duration: " + info.durationSeconds + "s\n" +
          "REMOTION_DATA: " + (info.remotionData || "")
        : "INFOGRAPHIC BUILD (" + info.tool + ")\n" +
          "Duration: " + info.durationSeconds + "s\n" +
          "Visual form: " + info.visualForm + "\n\n" +
          info.capCutInstructions;

      // ── Write the new infographic scene row ────────────────────────────────
      visualSheet.getRange(writeRow, COL_VISUAL.ID            ).setValue(idea.id);
      visualSheet.getRange(writeRow, COL_VISUAL.SCENE_NUM     ).setValue("I" + info.num);
      visualSheet.getRange(writeRow, COL_VISUAL.SCENE_TYPE    ).setValue(sceneType);
      visualSheet.getRange(writeRow, COL_VISUAL.DESCRIPTION   ).setValue(description);
      visualSheet.getRange(writeRow, COL_VISUAL.SOURCE        ).setValue(info.tool);
      visualSheet.getRange(writeRow, COL_VISUAL.LINK          ).setValue("Build per Director Report — see Assembly Notes");
      visualSheet.getRange(writeRow, COL_VISUAL.LICENSE       ).setValue("Original");
      visualSheet.getRange(writeRow, COL_VISUAL.STATUS        ).setValue("Needed");
      visualSheet.getRange(writeRow, COL_VISUAL.BUILT_WHERE   ).setValue(isRemotionScene ? "Remotion — Programmatic" : info.tool + " — Infographic");
      visualSheet.getRange(writeRow, COL_VISUAL.ASSEMBLY_NOTES).setValue(assemblyNotes);
      visualSheet.getRange(writeRow, COL_VISUAL.VEO_PROMPT    ).setValue("N/A — built as data graphic");

      // ── Write Remotion-specific columns (19, 21) ───────────────────────────
      if (isRemotionScene && info.remotionData) {
        visualSheet.getRange(writeRow, COL_VISUAL_EXTENDED.REMOTION_DATA ).setValue(info.remotionData);
      }
      if (info.voiceoverSync) {
        visualSheet.getRange(writeRow, COL_VISUAL_EXTENDED.VOICEOVER_SYNC).setValue(info.voiceoverSync);
      }

      // Style the new row to distinguish it from regular scenes
      visualSheet.getRange(writeRow, 1, 1, 18)
        .setBackground("#FFF3CD")  // amber tint — infographic scenes stand out
        .setFontStyle("normal");

      added++;
      Logger.log("Director: added infographic scene I" + info.num + " after scene " + targetSceneNum);

    } catch (err) {
      Logger.log("Director: failed to write infographic " + info.num + ": " + err.message);
    }
  });

  return added;
}


// ══════════════════════════════════════════════════════════════════════════════
// EXPORT DIRECTOR REPORT TO GOOGLE DRIVE
// Creates a formatted Google Doc with the full Director analysis
// ══════════════════════════════════════════════════════════════════════════════
function exportDirectorReport(idea, master, rawOutput, directorReport, infographicsAdded) {

  const divider = "═".repeat(60);
  const line    = "─".repeat(60);
  const isArabic = idea.language === "Arabic" || idea.language === "Bilingual";

  let doc = "";

  // ── Header ─────────────────────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "GOVERNX — DIRECTOR REPORT\n";
  doc += divider + "\n";
  doc += "Content ID    : " + idea.id + "\n";
  doc += "Title         : " + (master ? master.title : idea.company) + "\n";
  doc += "Company       : " + idea.company + "\n";
  doc += "Discipline    : " + (master ? master.discipline : "") + "\n";
  doc += "Language      : " + idea.language + "\n";
  doc += "Generated     : " + new Date().toLocaleString() + "\n";
  doc += divider + "\n\n";

  // ── Overall verdict ─────────────────────────────────────────────────────────
  doc += "OVERALL VERDICT: " + directorReport.overallVerdict + "\n";
  doc += "REASON: " + directorReport.verdictReason + "\n\n";

  doc += "SUMMARY:\n";
  doc += "  Visual sync failures : " + directorReport.syncFailures   + "\n";
  doc += "  Infographics added   : " + infographicsAdded              + "\n";
  doc += "  Pacing sections weak : " + directorReport.pacingWeakCount + "\n\n";

  // ── Section A: Visual-Narrative Sync ────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION A — VISUAL-NARRATIVE SYNC\n";
  doc += line + "\n\n";

  directorReport.syncScenes.forEach(scene => {
    const icon = scene.syncVerdict === "PASS"  ? "✅" :
                 scene.syncVerdict === "WEAK"  ? "⚠️" : "⛔";
    doc += icon + " SCENE " + scene.sceneNum + " (" + scene.sceneType + ") — " + scene.syncVerdict + "\n";
    doc += "   " + scene.syncReason + "\n";
    if (scene.replacementBrief && scene.replacementBrief.trim() !== "") {
      doc += "   REPLACEMENT: " + scene.replacementBrief + "\n";
    }
    doc += "\n";
  });

  // ── Section B: Infographic Briefs ───────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION B — INFOGRAPHIC & DATA VISUALIZATION BRIEFS\n";
  doc += line + "\n\n";

  if (directorReport.infographics.length === 0) {
    doc += "No infographics required for this video.\n\n";
  } else {
    directorReport.infographics.forEach((info, idx) => {
      doc += "INFOGRAPHIC " + (idx + 1) + " — " + info.visualForm.toUpperCase() + "\n";
      doc += line.substring(0, 40) + "\n";
      doc += "Trigger moment : " + info.triggerMoment         + "\n";
      doc += "Voiceover sync : " + info.voiceoverSync         + "\n";
      doc += "Data required  : " + info.dataRequired          + "\n";
      doc += "Visual form    : " + info.visualForm            + "\n";
      doc += "Build tool     : " + info.tool                  + "\n";
      doc += "Duration       : " + info.durationSeconds + "s" + "\n";
      doc += "Position       : After scene " + info.insertAfterScene + " (" + info.action + ")\n\n";
      doc += "CAPCUT BUILD INSTRUCTIONS:\n";
      doc += info.capCutInstructions + "\n\n";
    });
  }

  // ── Section C: Pacing Audit ─────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION C — EMOTIONAL PACING AUDIT\n";
  doc += line + "\n\n";

  directorReport.pacingSections.forEach(section => {
    const icon = section.score === "STRONG"   ? "💪" :
                 section.score === "ADEQUATE" ? "✅" : "⚠️";
    doc += icon + " " + section.name + " — " + section.score + "\n";
    doc += "   Assessment: " + section.assessment + "\n";
    if (section.fix && section.fix.trim() !== "" && section.score === "WEAK") {
      doc += "   FIX: " + section.fix + "\n";
    }
    doc += "\n";
  });

  // ── Section D: Final Notes ─────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "SECTION D — DIRECTOR'S FINAL NOTES\n";
  doc += line + "\n\n";
  doc += directorReport.finalNotes + "\n\n";

  // ── Production checklist ───────────────────────────────────────────────────
  doc += divider + "\n";
  doc += "PRODUCTION CHECKLIST — BEFORE STAGE 9 (ASSEMBLY GUIDE)\n";
  doc += line + "\n\n";

  const failedScenes = directorReport.syncScenes.filter(s => s.syncVerdict === "FAIL" || s.syncVerdict === "WEAK");
  if (failedScenes.length > 0) {
    doc += "⛔ SCENES TO REPLACE OR FIX:\n";
    failedScenes.forEach(s => {
      doc += "  • Scene " + s.sceneNum + " (" + s.syncVerdict + "): " + s.syncReason.substring(0, 80) + "\n";
    });
    doc += "\n";
  }

  if (directorReport.infographics.length > 0) {
    doc += "📊 INFOGRAPHICS TO BUILD:\n";
    directorReport.infographics.forEach((info, idx) => {
      doc += "  • Infographic " + (idx + 1) + ": " + info.visualForm + " | Tool: " + info.tool +
             " | After scene " + info.insertAfterScene + "\n";
    });
    doc += "  NOTE: Infographic scenes (I1, I2...) have been added to Visual Library\n";
    doc += "        with amber background. Build them before running Stage 9.\n\n";
  }

  const weakPacing = directorReport.pacingSections.filter(p => p.score === "WEAK");
  if (weakPacing.length > 0) {
    doc += "⚡ PACING FIXES NEEDED:\n";
    weakPacing.forEach(p => {
      doc += "  • " + p.name + ": " + (p.fix || p.assessment).substring(0, 100) + "\n";
    });
    doc += "\n";
  }

  if (failedScenes.length === 0 && directorReport.infographics.length === 0 && weakPacing.length === 0) {
    doc += "✅ All checks passed. Proceed directly to Stage 9 — Assembly Guide.\n\n";
  }

  doc += divider + "\n";
  doc += "END OF DIRECTOR REPORT — " + idea.id + "\n";
  doc += divider + "\n";

  // ── Save to Google Drive ───────────────────────────────────────────────────
  const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
  const docTitle      = idea.id + " — " + idea.company + " — Director Report";

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

  Logger.log("Director Report saved: " + newDoc.getUrl());
  return newDoc.getUrl();
}


// ══════════════════════════════════════════════════════════════════════════════
// WRITE DIRECTOR REPORT LINK TO PUBLISHING TRACKER
// Writes the report link into the Notes column so it's always one click away
// ══════════════════════════════════════════════════════════════════════════════
function writeDirectorLinkToPublishing(contentId, company, master, reportUrl) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const pubSheet = ss.getSheetByName(SHEET.PUBLISHING);
  if (!pubSheet) return;

  const pubData = pubSheet.getDataRange().getValues();

  for (let i = 1; i < pubData.length; i++) {
    if (pubData[i][COL_PUBLISHING.ID - 1].toString().trim() === contentId) {
      const currentNotes = pubData[i][COL_PUBLISHING.NOTES - 1].toString();
      const directorNote = "\n🎬 Director Report: " + reportUrl;

      // Append to existing notes — never overwrite
      pubSheet.getRange(i + 1, COL_PUBLISHING.NOTES)
        .setValue(currentNotes + directorNote);
      return;
    }
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// MENU INTEGRATION
// Add to GovernX menu between Stage 4 and Stage 5
// ══════════════════════════════════════════════════════════════════════════════
// HOW TO ADD TO MENU (edit Menu.gs):
//
// In Menu.gs, inside onOpen(), add this line AFTER Stage 4 and BEFORE Stage 5:
//
//   .addItem("🎬 Stage 4.5 — Director Pass",              "runDirectorPass")
//
// Full context:
//   .addItem("4️⃣  Stage 4 — Generate Scenes",              "generateScenes")
//   .addItem("🎬 Stage 4.5 — Director Pass",              "runDirectorPass")   ← ADD THIS
//   .addItem("5️⃣  Stage 5 — Create Publishing Row",        "createPublishingRow")
//
// ══════════════════════════════════════════════════════════════════════════════