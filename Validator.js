/* ============================================================================
   validator.gs — GovernX Content OS
   Quality gate checker — validates Claude output before writing to sheet
   Parses the QUALITY_CHECK block and enforces all standards
   ============================================================================ */

// ── Main validator — called after every Claude response ──────────────────────
function validateOutput(stage, rawOutput, contentId, targetFormat) {

  const result = {
    passed   : true,
    failures : [],
    warnings : [],
    qcData   : {}
  };

  // ── Stage-specific validation ───────────────────────────────────────────────
  switch (stage) {
    case "MASTER"   : validateMasterOutput(rawOutput, result);   break;
    case "RESEARCH" : validateResearchOutput(rawOutput, result); break;
    case "SCRIPT"   : validateScriptOutput(rawOutput, result, targetFormat, contentId); break;
    case "VISUAL"   : validateVisualOutput(rawOutput, result);   break;
    default:
      result.warnings.push("Unknown stage: " + stage + " — skipping validation");
  }

  return result;
}

// ── Master Content Validator ──────────────────────────────────────────────────
function validateMasterOutput(output, result) {

  const required = [
    "TITLE:", "DOMAIN:", "INDUSTRY:", "FIELD:", "TYPE:",
    "HOOK:", "FINAL_MOMENT:", "REVERSE_ANGLE:", "PRIMARY_ANGLE:",
    "DISCIPLINE:", "CORE_INSIGHT:", "CHECKPOINTS:",
    "TARGET_AUDIENCE:", "SERIES:"
  ];

  required.forEach(field => {
    if (!output.includes(field)) {
      result.failures.push("Missing required field: " + field);
      result.passed = false;
    }
  });

  // Discipline must be one of the valid values
  const disciplineMatch = output.match(/DISCIPLINE:\s*(.+)/);
  if (disciplineMatch) {
    const disc = disciplineMatch[1].trim();
    if (!["GRC", "BPR", "GRC+BPR"].includes(disc)) {
      result.failures.push(
        "Invalid DISCIPLINE value: '" + disc + "'. Must be GRC, BPR, or GRC+BPR"
      );
      result.passed = false;
    }
  }
}

// ── Research Database Validator ───────────────────────────────────────────────
function validateResearchOutput(output, result) {

  // Count number of sources returned
  const sourceMatches = output.match(/SOURCE_\d+_START/g);
  const sourceCount   = sourceMatches ? sourceMatches.length : 0;

  if (sourceCount < 6) {
    result.failures.push(
      "Insufficient sources: " + sourceCount + " found, minimum 6 required"
    );
    result.passed = false;
  }

  // Check source type diversity
  const sourceTypes = [];
  const typeMatches = output.matchAll(/SOURCE_TYPE:\s*(.+)/g);
  for (const match of typeMatches) {
    sourceTypes.push(match[1].trim());
  }

  const uniqueTypes = [...new Set(sourceTypes)];
  if (uniqueTypes.length < 2) {
    result.failures.push(
      "Insufficient source diversity: only 1 source type found. Minimum 2 required."
    );
    result.passed = false;
  }

  // Check minimum High relevance sources
  const highMatches  = output.match(/RELEVANCE:\s*High/gi);
  const highCount    = highMatches ? highMatches.length : 0;
  if (highCount < 3) {
    result.failures.push(
      "Insufficient high-relevance sources: " + highCount + " found, minimum 3 required"
    );
    result.passed = false;
  }

  // Check no empty KEY_INSIGHT fields
  const insightMatches = output.matchAll(/KEY_INSIGHT:\s*(.+)/g);
  let emptyInsights    = 0;
  for (const match of insightMatches) {
    if (!match[1] || match[1].trim() === "") emptyInsights++;
  }
  if (emptyInsights > 0) {
    result.failures.push(
      emptyInsights + " source(s) have empty KEY_INSIGHT. All sources must have a key insight."
    );
    result.passed = false;
  }
}

// ── Script Bank Validator ─────────────────────────────────────────────────────
function validateScriptOutput(output, result, targetFormat, contentId) {

  // ── Parse the QUALITY_CHECK block ────────────────────────────────────────
  const qcBlock = output.match(
    /QUALITY_CHECK_START([\s\S]*?)QUALITY_CHECK_END/
  );

  if (!qcBlock) {
    result.failures.push(
      "QUALITY_CHECK block missing from Claude output. Cannot validate script."
    );
    result.passed = false;
    return;
  }

  const qcText = qcBlock[1];
  const qc     = parseQualityCheck(qcText);
  result.qcData = qc;

  // ── Gate 1: Reverse engineering structure ────────────────────────────────
  if (qc.reverse_engineering_structure !== "YES") {
    result.failures.push(
      "Script does not follow GovernX reverse engineering structure " +
      "(Outcome → Decision Chain → Root Cause → Lesson)"
    );
    result.passed = false;
  }

  // ── Gate 2: GRC/BPR closing argument ────────────────────────────────────
  if (qc.grc_bpr_closing_argument !== "YES") {
    result.failures.push(
      "Script is missing the GRC/BPR closing argument. " +
      "Every GovernX video must name the governance or process discipline explicitly."
    );
    result.passed = false;
  }

  // ── Gate 3: Hook opens with specific fact ────────────────────────────────
  if (qc.hook_opens_with_specific_fact !== "YES") {
    result.failures.push(
      "Hook does not open with a specific date, number, or event. " +
      "GovernX hooks must create urgency with a concrete fact in the first 5 seconds."
    );
    result.passed = false;
  }

  // ── Gate 4: Minimum source references ────────────────────────────────────
  const sourcesReferenced = parseInt(qc.sources_referenced, 10) || 0;
  if (sourcesReferenced < 4) {
    result.failures.push(
      "Script references only " + sourcesReferenced + " sources. " +
      "Minimum 4 sources required from the Research Database."
    );
    result.passed = false;
  }

  // ── Gate 5: Word count range ──────────────────────────────────────────────
  if (targetFormat && WORD_COUNT[targetFormat]) {
    const wordCount = parseInt(qc.word_count, 10) || 0;
    const range     = WORD_COUNT[targetFormat];
    if (wordCount < range.min || wordCount > range.max) {
      result.warnings.push(
        "Word count " + wordCount + " is outside the recommended range " +
        "(" + range.min + "–" + range.max + ") for format: " + targetFormat
      );
      // Warning only — does not block writing to sheet
    }
  }

  // ── Gate 6: Living persons flag ──────────────────────────────────────────
  if (qc.living_persons_flagged === "YES") {
    result.warnings.push(
      "Script contains references to living persons. " +
      "Verify all claims use cautious language (reported, according to, evidence suggests)."
    );
  }

  // ── Gate 7: Copyright risks ───────────────────────────────────────────────
  if (qc.copyright_risks_flagged === "YES") {
    result.warnings.push(
      "Script or scene list contains copyright risk flags. " +
      "Review [COPYRIGHT RISK] items before production."
    );
  }

  // ── Gate 8: Scene blueprint present ──────────────────────────────────────
  if (!output.includes("SCENE_BLUEPRINT_START")) {
    result.failures.push(
      "SCENE_BLUEPRINT missing. Every script must include a complete scene-by-scene " +
      "production plan so Stage 4 can execute it without interpretation."
    );
    result.passed = false;
  }

  // ── Gate 9: Minimum 4 data moments ───────────────────────────────────────
  const dmMatches = output.match(/DM_\d+:/g) || [];
  if (dmMatches.length < 4) {
    result.failures.push(
      "Insufficient DATA_MOMENTS: " + dmMatches.length + " found, minimum 4 required. " +
      "Every quantifiable moment in the voiceover must be listed with specific figures and years."
    );
    result.passed = false;
  }
}

// ── Visual Library Validator ──────────────────────────────────────────────────
function validateVisualOutput(output, result) {

  // Count scenes
  const sceneMatches = output.match(/SCENE_\d+_START/g);
  const sceneCount   = sceneMatches ? sceneMatches.length : 0;

  if (sceneCount < 3) {
    result.failures.push(
      "Insufficient scenes: " + sceneCount + " found, minimum 3 required " +
      "(1 opening + at least 1 checkpoint + 1 timeline)"
    );
    result.passed = false;
  }

  // Check that a Timeline scene exists as the last scene
  const hasTimeline = output.includes("SCENE_TYPE: Timeline");
  if (!hasTimeline) {
    result.failures.push(
      "Missing Timeline scene. Every GovernX video must end with a Timeline scene " +
      "showing all checkpoints together."
    );
    result.passed = false;
  }

  // Check that at least one Checkpoint scene exists
  const hasCheckpoint = output.includes("SCENE_TYPE: Checkpoint");
  if (!hasCheckpoint) {
    result.warnings.push(
      "No Checkpoint scenes detected. GovernX videos should include checkpoint cards " +
      "to show the reverse-engineering timeline."
    );
  }

  // Check no scene is missing a source
  const sourceMatches  = output.matchAll(/SCENE_SOURCE:\s*(.+)/g);
  let   missingSource  = 0;
  for (const match of sourceMatches) {
    if (!match[1] || match[1].trim() === "") missingSource++;
  }
  if (missingSource > 0) {
    result.failures.push(
      missingSource + " scene(s) are missing a source. Every scene must have a suggested source."
    );
    result.passed = false;
  }

  // Flag license risks
  const riskMatches = output.match(/LICENSE:\s*Risk/gi);
  if (riskMatches && riskMatches.length > 0) {
    result.warnings.push(
      riskMatches.length + " scene(s) have LICENSE: Risk. " +
      "Review these before production."
    );
  }
}

// ── Parse QUALITY_CHECK block into key-value object ───────────────────────────
function parseQualityCheck(qcText) {
  const qc    = {};
  const lines = qcText.trim().split("\n");
  lines.forEach(line => {
    const parts = line.split(":");
    if (parts.length >= 2) {
      const key   = parts[0].trim().toLowerCase().replace(/ /g, "_");
      const value = parts.slice(1).join(":").trim();
      qc[key]     = value;
    }
  });
  return qc;
}

// ── Show validation result to user ───────────────────────────────────────────
function showValidationResult(stage, result, contentId) {

  const ui = SpreadsheetApp.getUi();

  if (result.passed && result.warnings.length === 0) {
    // All good — no popup needed, just proceed
    return true;
  }

  if (!result.passed) {
    // Build failure message
    let msg = "⛔ STAGE " + stage + " — QUALITY GATE FAILED\n";
    msg    += "Content ID: " + contentId + "\n\n";
    msg    += "FAILURES (output NOT written to sheet):\n";
    result.failures.forEach((f, i) => { msg += (i + 1) + ". " + f + "\n"; });

    if (result.warnings.length > 0) {
      msg += "\nWARNINGS:\n";
      result.warnings.forEach((w, i) => { msg += (i + 1) + ". " + w + "\n"; });
    }

    msg += "\nCheck the Error Log tab for details.";
    ui.alert("Quality Gate Failed", msg, ui.ButtonSet.OK);
    return false;
  }

  if (result.warnings.length > 0) {
    // Warnings only — ask user whether to proceed
    let msg = "⚠️ STAGE " + stage + " — WARNINGS FOUND\n";
    msg    += "Content ID: " + contentId + "\n\n";
    msg    += "Output passed all quality gates but has warnings:\n\n";
    result.warnings.forEach((w, i) => { msg += (i + 1) + ". " + w + "\n"; });
    msg    += "\nProceed and write to sheet anyway?";

    const response = ui.alert(
      "Quality Warnings", msg, ui.ButtonSet.YES_NO
    );
    return response === ui.Button.YES;
  }

  return true;
}