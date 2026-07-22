/* ============================================================================
   Stage_8D_Remotion.gs — GovernX Content OS
   Stage 8D — Remotion Programmatic Scene Renderer
   ============================================================================ */


// ── Render server URL ─────────────────────────────────────────────────────────
function getRemotionServerUrl() {
  const url = PropertiesService.getScriptProperties().getProperty("REMOTION_SERVER_URL");
  if (!url || url.trim() === "") {
    throw new Error(
      "REMOTION_SERVER_URL not set in Script Properties.\n\n" +
      "Add it now:\n" +
      "1. Apps Script → Project Settings → Script Properties\n" +
      "2. Key: REMOTION_SERVER_URL\n" +
      "3. Value: https://skyrocket-overfeed-taunt.ngrok-free.dev"
    );
  }
  return url.trim().replace(/\/$/, "");
}

const REMOTION_SCENE_TYPES = ["Checkpoint", "Infographic", "Text", "Timeline", "Data Table", "Opening Title", "Title", "Risk Matrix", "KPI Dashboard", "Gauge"];


// ══════════════════════════════════════════════════════════════════════════════
// STAGE 8D — MAIN FUNCTION
// ══════════════════════════════════════════════════════════════════════════════
function renderRemotionScenes() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!checkRendererHealth(ui)) return;

  const master = getMasterContent(idea.id);

  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const scenes      = [];

  let checkpointTotal = 0;
  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    if (visualData[i][COL_VISUAL.SCENE_TYPE - 1] === "Checkpoint") checkpointTotal++;
  }

  let checkpointNum = 0;

  for (let i = 1; i < visualData.length; i++) {
    const row = visualData[i];
    if (row[COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;

    const sceneType = row[COL_VISUAL.SCENE_TYPE - 1].toString().trim();
    const status    = row[COL_VISUAL.STATUS     - 1].toString().trim();
    const sceneNum  = row[COL_VISUAL.SCENE_NUM  - 1].toString().trim();

    if (status === "Done" || status === "Skip") continue;

    if (!REMOTION_SCENE_TYPES.some(t =>
      sceneType.toLowerCase().includes(t.toLowerCase())
    )) continue;

    if (sceneType === "Checkpoint") checkpointNum++;

    scenes.push({
      row             : i + 1,
      sceneNum,
      sceneType,
      description     : row[COL_VISUAL.DESCRIPTION      - 1].toString(),
      assemblyNotes   : row[COL_VISUAL.ASSEMBLY_NOTES   - 1].toString(),
      checkpointDate  : row[COL_VISUAL.CHECKPOINT_DATE  - 1].toString().trim(),
      checkpointEvent : row[COL_VISUAL.CHECKPOINT_EVENT - 1].toString().trim(),
      checkpointAngle : row[COL_VISUAL.CHECKPOINT_ANGLE - 1].toString().trim(),
      checkpointNum   : checkpointNum,
      checkpointTotal : checkpointTotal,
      timestamp       : row[COL_VISUAL.TIMESTAMP - 1].toString().trim(),
      remotionData    : row.length > 18 ? row[COL_VISUAL_EXTENDED.REMOTION_DATA  - 1].toString().trim() : "",
      remotionStyle   : row.length > 19 ? row[COL_VISUAL_EXTENDED.REMOTION_STYLE - 1].toString().trim() : "",
      voiceoverSync   : row.length > 20 ? row[COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1].toString().trim() : "",
      sceneScore      : row.length > 21 ? row[COL_VISUAL_EXTENDED.SCENE_SCORE    - 1].toString().trim() : ""
    });
  }

  if (scenes.length === 0) {
    ui.alert(
      "No Remotion Scenes Found",
      "No Checkpoint, Infographic, Text, or Opening Title scenes found\n" +
      "with status 'Needed' or 'Submitted' for: " + idea.id + "\n\n" +
      "Either all scenes are already Done/Skip, or no Remotion scene\n" +
      "types exist in the Visual Library for this content ID.",
      ui.ButtonSet.OK
    );
    return;
  }

  const confirm = ui.alert(
    "🎬 Stage 8D — Remotion Renderer",
    "Found " + scenes.length + " scene(s) to render for: " + idea.company + "\n\n" +
    scenes.map(s => "  • Scene " + s.sceneNum + " — " + s.sceneType).join("\n") +
    "\n\nRender all scenes programmatically?\n" +
    "(Make sure governx-remotion server is running)",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  // ── Pre-flight: verify available compositions ─────────────────────────────
  // Checks /compositions endpoint on the Remotion server and warns if
  // TimelineReveal is missing before attempting any renders.
  try {
    const compResponse = UrlFetchApp.fetch(getRemotionServerUrl() + "/compositions", {
      muteHttpExceptions: true,
      headers: { "ngrok-skip-browser-warning": "true" }
    });
    if (compResponse.getResponseCode() === 200) {
      const compBody  = JSON.parse(compResponse.getContentText());
      const available = Array.isArray(compBody)
        ? compBody.map(c => c.id || c)
        : (compBody.compositions || []);
      const required  = ["CheckpointCard", "InfographicScene", "TextImpactScene", "TimelineReveal", "OpeningTitle", "RiskMatrix", "KPIDashboard", "ProgressGauge"];
      const missing   = required.filter(r => !available.includes(r));
      if (missing.length > 0) {
        const hasTimeline = scenes.some(s => s.sceneType === "Timeline");
        if (missing.includes("TimelineReveal") && hasTimeline) {
          const proceed = ui.alert(
            "⚠️ TimelineReveal Composition Missing",
            "Your Remotion project is missing the 'TimelineReveal' composition.\n\n" +
            "This video has " + scenes.filter(s => s.sceneType === "Timeline").length +
            " Timeline scene(s) that will fail to render.\n\n" +
            "To fix: open your Remotion project, create a TimelineReveal composition\n" +
            "in src/Root.jsx, then restart the Remotion server.\n\n" +
            "Available now: " + (available.length ? available.join(", ") : "none detected") + "\n\n" +
            "Continue anyway and skip Timeline scenes?",
            ui.ButtonSet.YES_NO
          );
          if (proceed !== ui.Button.YES) return;
        }
      }
    }
  } catch (e) {
    Logger.log("Stage 8D: Could not verify compositions (non-fatal): " + e.message);
  }

  const renderPayloads = scenes.map(scene => buildRenderPayload(scene, idea, master, scenes));

  // ── REMOTION_DATA validation — catch bad data before hitting the server ──────
  const validationIssues = validateRemotionPayloads_(scenes, renderPayloads);
  if (validationIssues.length > 0) {
    const proceed = ui.alert(
      "⚠️ Stage 8D — Data Validation Issues",
      validationIssues.length + " scene(s) have missing or invalid data:\n\n" +
      validationIssues.map(v => "• Scene " + v.sceneNum + " (" + v.type + "): " + v.issue).join("\n") +
      "\n\nThese scenes will likely render blank or with placeholder content.\n\n" +
      "Fix the REMOTION_DATA in the Visual Library, OR continue anyway?",
      ui.ButtonSet.YES_NO
    );
    if (proceed !== ui.Button.YES) return;
  }

  const contentFolder = getOrCreateContentFolder(idea.id, idea.company);
  const scenesFolder  = getOrCreateScenesFolder(idea.id, idea.company);

  let succeeded = 0;
  let failed    = 0;
  const results = [];

  scenes.forEach((scene, idx) => {
    const payload = renderPayloads[idx];
    if (!payload) return;

    try {
      visualSheet.getRange(scene.row, COL_VISUAL.STATUS).setValue("Submitted");
      SpreadsheetApp.flush();

      const response = UrlFetchApp.fetch(getRemotionServerUrl() + "/render", {
        method            : "post",
        contentType       : "application/json",
        payload           : JSON.stringify(payload),
        muteHttpExceptions: true,
        followRedirects   : true
      });

      const code = response.getResponseCode();
      const body = JSON.parse(response.getContentText());

      if (code === 200 && body.success) {
        const serverBase  = getRemotionServerUrl();
        const mp4Url      = body.url.replace(/https?:\/\/localhost:\d+/, serverBase)
                                     .replace(/https?:\/\/127\.0\.0\.1:\d+/, serverBase);
        Logger.log("Stage 8D: Downloading MP4 from: " + mp4Url);
        const mp4Response = UrlFetchApp.fetch(mp4Url, {
          muteHttpExceptions: true,
          headers: { "ngrok-skip-browser-warning": "true" }
        });

        const mp4ContentType = mp4Response.getHeaders()["content-type"] || "";
        if (mp4Response.getResponseCode() === 200 &&
            !mp4ContentType.toLowerCase().includes("text") &&
            !mp4ContentType.toLowerCase().includes("html")) {
          const blob      = mp4Response.getBlob().setContentType("video/mp4").setName(body.filename);
          const driveFile = scenesFolder.createFile(blob);
          const driveUrl  = driveFile.getUrl();

          visualSheet.getRange(scene.row, COL_VISUAL.STATUS        ).setValue("Done");
          visualSheet.getRange(scene.row, COL_VISUAL.AI_CLIP_URL   ).setValue(driveUrl);
          visualSheet.getRange(scene.row, COL_VISUAL.BUILT_WHERE   ).setValue("Remotion — Local");
          visualSheet.getRange(scene.row, COL_VISUAL.ASSEMBLY_NOTES).setValue(
            (scene.assemblyNotes ? scene.assemblyNotes + "\n" : "") +
            "Remotion MP4: " + driveUrl
          );
          visualSheet.getRange(scene.row, 1, 1, 18).setBackground("#E8F5E9");

          succeeded++;
          results.push({ sceneNum: scene.sceneNum, success: true, driveUrl });
          Logger.log("Stage 8D: Scene " + scene.sceneNum + " rendered → " + driveUrl);
        } else {
          const ct = mp4ContentType || "unknown";
          throw new Error(
            mp4Response.getResponseCode() !== 200
              ? "Could not download rendered file from server (HTTP " + mp4Response.getResponseCode() + ")"
              : "Server returned " + ct + " instead of an MP4.\n" +
                "If using ngrok, make sure it is running and the REMOTION_SERVER_URL is current.\n" +
                "Then retry Stage 8D."
          );
        }

      } else {
        throw new Error(body.error || "Render failed with code " + code);
      }

    } catch (err) {
      failed++;
      visualSheet.getRange(scene.row, COL_VISUAL.STATUS).setValue("Error");
      visualSheet.getRange(scene.row, COL_VISUAL.NOTE  ).setValue("8D Error: " + err.message);
      visualSheet.getRange(scene.row, 1, 1, 18).setBackground("#FFE0E0");
      logError("Stage 8D — Remotion", idea.id, "Render Error Scene " + scene.sceneNum, err.message);
      results.push({ sceneNum: scene.sceneNum, success: false, error: err.message });
    }

    SpreadsheetApp.flush();
  });

  updatePipelineStatus_(idea.id, "S8D", failed === 0 ? "✅" : (succeeded > 0 ? "⚠️" : "❌"));
  ui.alert(
    succeeded === scenes.length ? "✅ Stage 8D Complete" : "⚠️ Stage 8D Partial",
    "Rendered: " + succeeded + "/" + scenes.length + " scenes\n" +
    (failed > 0 ? "Failed : " + failed + " (see Error Log + red rows in Visual Library)\n\n" : "\n") +
    "Successful renders saved to Drive folder:\n" + idea.company + "\n\n" +
    (succeeded > 0 ? "Proceed to Stage 9 — Assembly Guide." : "Fix errors and re-run Stage 8D."),
    ui.ButtonSet.OK
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// BUILD RENDER PAYLOAD
// ══════════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════════
// BUILD RENDER PAYLOAD
// Selects composition variant based on idea.targetFormat:
//   Short (<60s) / Short (< 90s)  → _Vertical compositions (1080×1920)
//   Standard / Deep Dive          → base compositions     (1920×1080)
// ══════════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════════
// BUILD RENDER PAYLOAD
// Selects composition variant based on idea.targetFormat:
//   Short (<60s) / Short (< 90s)  → -Vertical compositions (1080×1920)
//   Standard / Deep Dive          → base compositions     (1920×1080)
// ══════════════════════════════════════════════════════════════════════════════
function buildRenderPayload(scene, idea, master, allScenes) {

  const baseFilename = idea.id + "_scene_" + scene.sceneNum;

  // ── Format-aware composition suffix ─────────────────────────────────────────
  const isShort = idea.targetFormat
    ? idea.targetFormat.toLowerCase().includes("short")
    : false;
  const suffix = isShort ? "-Vertical" : "";

  // ── CHECKPOINT CARD ─────────────────────────────────────────────────────────
  if (scene.sceneType === "Checkpoint") {
    if (!scene.checkpointDate || !scene.checkpointEvent) {
      Logger.log("Stage 8D: Skipping checkpoint scene " + scene.sceneNum + " — missing date/event");
      return null;
    }
    const styleVariant = parseStyleKey(scene.remotionStyle, "variant") || "standard";
    return {
      compositionId   : "CheckpointCard" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_checkpoint.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 5),
      props: {
        date             : cleanCheckpointDate(scene.checkpointDate),
        event            : scene.checkpointEvent,
        angle            : scene.checkpointAngle || "GOVERNANCE FAILURE",
        checkpointNum    : scene.checkpointNum,
        totalCheckpoints : scene.checkpointTotal,
        variant          : styleVariant
      }
    };
  }

  // ── INFOGRAPHIC SCENE ────────────────────────────────────────────────────────
  if (scene.sceneType === "Infographic") {
    // PRIMARY PATH: REMOTION_DATA column (col 19) — Stage 4 Director pass output
    if (scene.remotionData && scene.remotionData.length > 10) {
      const rd      = parseRemotionData(scene.remotionData);
      const variant = parseStyleKey(scene.remotionStyle, "variant") || rd.type || "DATA_CALLOUT";
      return {
        compositionId   : "InfographicScene" + suffix,
        contentId       : idea.id,
        sceneNum        : scene.sceneNum,
        outputFilename  : baseFilename + "_infographic.mp4",
        durationInFrames: calcSceneDuration(scene.voiceoverSync, 6),
        props: buildInfographicPropsFromData(rd, variant, scene, idea)
      };
    }
    // FALLBACK: parse from ASSEMBLY_NOTES / DESCRIPTION (handles pipe format too)
    const parsed = parseInfographicBrief(scene.assemblyNotes || scene.description);
    const inferredType = parsed.type || "data_callout";
    const props = {
      type          : inferredType,
      title         : parsed.title    || idea.company,
      subtitle      : parsed.subtitle || "",
      dataPoints    : parsed.dataPoints || [],
      unit          : parsed.unit       || "",
      voiceoverSync : scene.voiceoverSync || parsed.voiceoverSync || ""
    };
    if (inferredType === "split_comparison") {
      props.leftLabel   = parsed.leftLabel   || "";
      props.leftValue   = parsed.leftValue   || "";
      props.rightLabel  = parsed.rightLabel  || "";
      props.rightValue  = parsed.rightValue  || "";
      props.footer      = parsed.footer      || "";
      props.colorScheme = parsed.colorScheme || "green_to_red";
    }
    return {
      compositionId   : "InfographicScene" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_infographic.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 6),
      props           : props
    };
  }

  // ── TEXT IMPACT SCENE ────────────────────────────────────────────────────────
  if (scene.sceneType.toLowerCase().includes("text")) {
    if (scene.remotionData && scene.remotionData.length > 10) {
      const rd      = parseRemotionData(scene.remotionData);
      const variant = parseStyleKey(scene.remotionStyle, "variant") || rd.type || "default";
      return {
        compositionId   : "TextImpactScene" + suffix,
        contentId       : idea.id,
        sceneNum        : scene.sceneNum,
        outputFilename  : baseFilename + "_text.mp4",
        durationInFrames: calcSceneDuration(scene.voiceoverSync, 5),
        props: {
          type    : variant,
          mainText: rd.mainText || "",
          subText : rd.subText  || "",
          context : idea.company,
          accent  : true
        }
      };
    }
    const parsed = parseTextBrief(scene.description || scene.assemblyNotes);
    return {
      compositionId   : "TextImpactScene" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_text.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 5),
      props: {
        type    : parsed.type     || "stat",
        mainText: parsed.mainText || scene.description.substring(0, 80),
        subText : parsed.subText  || "",
        context : parsed.context  || idea.company,
        accent  : parsed.accent !== false
      }
    };
  }

  // ── OPENING TITLE ────────────────────────────────────────────────────────────
  if (scene.sceneType.toLowerCase().includes("title") ||
      scene.sceneType.toLowerCase().includes("opening")) {
    return {
      compositionId   : "OpeningTitle" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_title.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 4),
      props: {
        company   : idea.company,
        discipline: master ? master.discipline || "GRC" : "GRC",
        hook      : master ? master.hook || "" : "",
        contentId : idea.id
      }
    };
  }

  // ── TIMELINE SCENE ───────────────────────────────────────────────────────────
  if (scene.sceneType === "Timeline") {
    const rd = scene.remotionData ? parseRemotionData(scene.remotionData) : {};

    let checkpointStrings = [];

    if (allScenes && allScenes.length > 0) {
      const cpScenes = allScenes.filter(s => s.sceneType === "Checkpoint");
      if (cpScenes.length > 0) {
        checkpointStrings = cpScenes.map(cp => {
          const date  = cleanCheckpointDate(cp.checkpointDate) || "";
          const event = (cp.checkpointEvent || cp.description || "").substring(0, 70).trim();
          return event ? event + (date ? " " + date : "") : date;
        }).filter(Boolean);
      }
    }

    if (checkpointStrings.length === 0 && rd.checkpoints) {
      checkpointStrings = rd.checkpoints.split(",").map(c => c.trim()).filter(Boolean);
    }

    if (checkpointStrings.length === 0 && scene.assemblyNotes) {
      const cpMatch = scene.assemblyNotes.match(/checkpoints=\[?([^\]|]+)/i);
      if (cpMatch) {
        checkpointStrings = cpMatch[1].split(",").map(c => c.trim()).filter(Boolean);
      }
    }

    return {
      compositionId   : "TimelineReveal" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_timeline.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 10),
      props: {
        checkpoints   : checkpointStrings,
        company       : idea.company,
        voiceoverSync : scene.voiceoverSync || ""
      }
    };
  }

  // ── RISK MATRIX ──────────────────────────────────────────────────────────────
  // REMOTION_DATA format:
  //   type=RISK_MATRIX | title=My Title | xLabel=LIKELIHOOD | yLabel=IMPACT |
  //   risks=Label One:3:3:true,Label Two:2:1:false,...
  //   (each risk: label:likelihood:impact:highlight)
  if (scene.sceneType === "Risk Matrix") {
    const rd = scene.remotionData ? parseRemotionData(scene.remotionData) : {};
    const risks = (rd.risks || "").split(",").map(r => {
      const parts = r.trim().split(":");
      return {
        label      : parts[0] ? parts[0].trim() : "",
        likelihood : parseInt(parts[1]) || 1,
        impact     : parseInt(parts[2]) || 1,
        highlight  : (parts[3] || "").trim().toLowerCase() === "true"
      };
    }).filter(r => r.label);
    return {
      compositionId   : "RiskMatrix" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_riskmatrix.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 8),
      props: {
        title  : rd.title  || idea.company + " — Risk Assessment",
        xLabel : rd.xlabel || "LIKELIHOOD",
        yLabel : rd.ylabel || "IMPACT",
        risks  : risks.length > 0 ? risks : []
      }
    };
  }

  // ── KPI DASHBOARD ─────────────────────────────────────────────────────────────
  // REMOTION_DATA format:
  //   type=KPI_DASHBOARD | title=My Title | layout=2x2 |
  //   kpis=Label:Value:trend:change:context:highlight,...
  //   (each kpi: label:value:up|down|neutral:change label:context:true|false)
  if (scene.sceneType === "KPI Dashboard") {
    const rd = scene.remotionData ? parseRemotionData(scene.remotionData) : {};
    const kpis = (rd.kpis || "").split(",").map(k => {
      const parts = k.trim().split(":");
      return {
        label    : parts[0] ? parts[0].trim() : "",
        value    : parts[1] ? parts[1].trim() : "",
        trend    : parts[2] ? parts[2].trim().toLowerCase() : "neutral",
        change   : parts[3] ? parts[3].trim() : "",
        context  : parts[4] ? parts[4].trim() : "",
        highlight: (parts[5] || "").trim().toLowerCase() === "true"
      };
    }).filter(k => k.label && k.value);
    return {
      compositionId   : "KPIDashboard" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_kpi.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 7),
      props: {
        title  : rd.title  || idea.company + " — KPI Dashboard",
        layout : rd.layout || "2x2",
        kpis   : kpis.length > 0 ? kpis : []
      }
    };
  }

  // ── PROGRESS GAUGE ────────────────────────────────────────────────────────────
  // REMOTION_DATA format:
  //   type=PROGRESS_GAUGE | title=My Title | variant=multi |
  //   gauges=Label:value:unit:context:highlight:threshold,...
  //   (each gauge: label:0-100:unit:%|/100|etc:context:true|false:threshold 0-100)
  if (scene.sceneType === "Gauge") {
    const rd = scene.remotionData ? parseRemotionData(scene.remotionData) : {};
    const gauges = (rd.gauges || "").split(",").map(g => {
      const parts = g.trim().split(":");
      const gauge = {
        label    : parts[0] ? parts[0].trim() : "",
        value    : parseFloat(parts[1]) || 0,
        unit     : parts[2] ? parts[2].trim() : "%",
        context  : parts[3] ? parts[3].trim() : "",
        highlight: (parts[4] || "").trim().toLowerCase() === "true"
      };
      if (parts[5] && parts[5].trim() !== "") gauge.threshold = parseFloat(parts[5].trim());
      return gauge;
    }).filter(g => g.label);
    return {
      compositionId   : "ProgressGauge" + suffix,
      contentId       : idea.id,
      sceneNum        : scene.sceneNum,
      outputFilename  : baseFilename + "_gauge.mp4",
      durationInFrames: calcSceneDuration(scene.voiceoverSync, 7),
      props: {
        title  : rd.title   || idea.company + " — Compliance",
        variant: rd.variant || (gauges.length > 1 ? "multi" : "single"),
        gauges : gauges.length > 0 ? gauges : []
      }
    };
  }

  // ── DATA TABLE SCENE ─────────────────────────────────────────────────────────
  if (scene.sceneType === "Data Table") {
    if (scene.remotionData && scene.remotionData.length > 10) {
      const rd      = parseRemotionData(scene.remotionData);
      const variant = parseStyleKey(scene.remotionStyle, "variant") || "BEFORE_AFTER_CARD";
      return {
        compositionId   : "InfographicScene" + suffix,
        contentId       : idea.id,
        sceneNum        : scene.sceneNum,
        outputFilename  : baseFilename + "_datatable.mp4",
        durationInFrames: calcSceneDuration(scene.voiceoverSync, 6),
        props: buildInfographicPropsFromData(rd, variant, scene, idea)
      };
    }
    return null;
  }

  Logger.log("Stage 8D: Unknown scene type '" + scene.sceneType + "' for scene " + scene.sceneNum);
  return null;
}


// ══════════════════════════════════════════════════════════════════════════════
// PARSE INFOGRAPHIC BRIEF
// Fallback parser used when REMOTION_DATA (col 19) is empty.
// Handles TWO formats written into ASSEMBLY_NOTES / DESCRIPTION:
//
// FORMAT A — pipe key=value (Stage 4B Director output, most common):
//   type=SPLIT_COMPARISON | left_label=X | left_value=Y | right_label=A |
//   right_value=B | footer=Z | title=CUSTOM | color_scheme=green_to_red
//
// FORMAT B — legacy labelled lines (older output):
//   [INFOGRAPHIC] SPLIT_COMPARISON — TITLE | Visual form: split_comparison |
//   Data required: Label: value | Voiceover sync: "..."
//
// FORMAT A is detected automatically when "type=" or "left_label=" is present.
// ══════════════════════════════════════════════════════════════════════════════
function parseInfographicBrief(text) {

  if (!text) return {};

  const result = {
    type         : "data_callout",
    title        : "",
    subtitle     : "",
    dataPoints   : [],
    unit         : "",
    voiceoverSync: "",
    leftLabel    : "",
    leftValue    : "",
    rightLabel   : "",
    rightValue   : "",
    footer       : "",
    colorScheme  : "green_to_red"
  };

  const typeMap = {
    "bar_chart"        : "bar_chart",
    "line_graph"       : "bar_chart",
    "split_comparison" : "split_comparison",
    "counter_animation": "counter",
    "percentage_ring"  : "counter",
    "before_after"     : "before_after",
    "data_callout"     : "data_callout",
    "timeline_bar"     : "data_callout"
  };

  // ── FORMAT A: pipe key=value ─────────────────────────────────────────────────
  const isFormatA = /\btype\s*=/i.test(text) || /\bleft_label\s*=/i.test(text);

  if (isFormatA) {
    const pairs = text.split(/\s*\|\s*/);
    const kv = {};
    pairs.forEach(function(pair) {
      const eqIdx = pair.indexOf("=");
      if (eqIdx === -1) return;
      const key = pair.substring(0, eqIdx).trim().toLowerCase().replace(/\s+/g, "_");
      const val = pair.substring(eqIdx + 1).trim();
      kv[key] = val;
    });

    const rawType = (kv["type"] || "").toLowerCase();
    result.type = typeMap[rawType] || "data_callout";

    result.leftLabel    = kv["left_label"]   || kv["left_title"]  || "";
    result.leftValue    = kv["left_value"]   || kv["left_stat"]   || "";
    result.rightLabel   = kv["right_label"]  || kv["right_title"] || "";
    result.rightValue   = kv["right_value"]  || kv["right_stat"]  || "";
    result.footer       = kv["footer"]       || kv["note"]        || "";
    result.colorScheme  = kv["color_scheme"] || kv["color"]       || "green_to_red";
    result.voiceoverSync = kv["voiceover_sync"] || kv["voiceover"] || "";

    if (kv["title"] && kv["title"].trim() !== "") {
      result.title = kv["title"].trim().substring(0, 80);
    } else if (result.leftLabel && result.rightLabel) {
      result.title = extractShortLabel(result.leftLabel) +
                     " VS " +
                     extractShortLabel(result.rightLabel);
    } else {
      result.title = kv["description"] || kv["subtitle"] || "";
    }

    result.subtitle = kv["subtitle"] || result.footer.substring(0, 60) || "";

    if (result.type === "split_comparison") {
      result.dataPoints = [
        { label: result.leftLabel,  value: result.leftValue,  highlight: false, side: "left"  },
        { label: result.rightLabel, value: result.rightValue, highlight: true,  side: "right" }
      ];
    }

    Logger.log("parseInfographicBrief [FORMAT A]: type=" + result.type + " title=" + result.title);
    return result;
  }

  // ── FORMAT B: legacy labelled-line format ───────────────────────────────────
  const formMatch = text.match(/Visual form\s*:\s*([a-z_]+)/i);
  if (formMatch) result.type = typeMap[formMatch[1].toLowerCase()] || "data_callout";

  const titleMatch = text.match(/\[INFOGRAPHIC\]\s+([A-Z_]+)\s+[—-]\s+(.+?)(?:\||$)/);
  if (titleMatch) {
    result.type  = typeMap[titleMatch[1].toLowerCase()] || result.type;
    result.title = titleMatch[2].trim().substring(0, 60);
  }

  const syncMatch = text.match(/Voiceover sync\s*:\s*(.+)/i);
  if (syncMatch) result.voiceoverSync = syncMatch[1].substring(0, 100).replace(/"/g, "");

  const dataMatch = text.match(/Data required\s*:\s*(.+)/i);
  if (dataMatch) {
    const dataStr    = dataMatch[1];
    const numPattern = /([A-Za-z][^:,]+?):\s*\$?([\d,.]+)([BMK%]?)/g;
    let match;
    let i = 0;
    while ((match = numPattern.exec(dataStr)) !== null && i < 8) {
      const value    = parseFloat(match[2].replace(",", ""));
      const unitHint = match[3] || "";
      if (!isNaN(value)) {
        result.dataPoints.push({
          label    : match[1].trim(),
          value    : value,
          highlight: i === result.dataPoints.length - 1
        });
        if (!result.unit && unitHint) {
          result.unit = unitHint === "B" ? "$B" : unitHint === "M" ? "$M" : unitHint;
        }
        i++;
      }
    }
    if (result.title) {
      result.subtitle = dataStr.replace(numPattern, "").trim().substring(0, 60);
    }
  }

  if (result.dataPoints.length === 0) {
    result.dataPoints = [{ label: result.title || "Data", value: 0, highlight: true }];
  }

  Logger.log("parseInfographicBrief [FORMAT B]: type=" + result.type + " title=" + result.title);
  return result;
}


// ── Extract a short display label from a long comparison label ───────────────
// "Pension at Acquisition (2000)" → "2000 — ACQUISITION"
// "Pension at Collapse (2016)"    → "2016 — COLLAPSE"
function extractShortLabel(label) {
  if (!label) return "";
  const yearMatch = label.match(/\((\d{4})\)/);
  const nounMatch = label.match(/\bat\s+([A-Za-z]+)/i);
  if (yearMatch && nounMatch) return yearMatch[1] + " — " + nounMatch[1].toUpperCase();
  if (yearMatch) return yearMatch[1];
  return label.toUpperCase().substring(0, 25).trim();
}


// ══════════════════════════════════════════════════════════════════════════════
// PARSE TEXT BRIEF
// ══════════════════════════════════════════════════════════════════════════════
function parseTextBrief(text) {

  if (!text) return {};

  const result = {
    type    : "stat",
    mainText: "",
    subText : "",
    context : "",
    accent  : true
  };

  const lower = text.toLowerCase();
  if (lower.includes("grc") || lower.includes("governance is") || lower.includes("closing")) {
    result.type = "grc_closing";
    result.accent = false;
  } else if (lower.includes("hook") || lower.includes("?")) {
    result.type = "hook";
    result.accent = false;
  } else if (lower.includes("verdict") || lower.includes("conclusion")) {
    result.type = "verdict";
    result.accent = true;
  } else {
    result.type = "stat";
    result.accent = true;
  }

  const quotedMatch = text.match(/"([^"]{10,120})"/);
  if (quotedMatch) {
    result.mainText = quotedMatch[1];
  } else {
    const lines = text.split("\n").map(l => l.trim()).filter(l =>
      l.length > 8 &&
      !l.startsWith("[") &&
      !l.startsWith("//") &&
      !l.match(/^[A-Z_]+:/)
    );
    result.mainText = lines[0] ? lines[0].substring(0, 100) : text.substring(0, 80);
    result.subText  = lines[1] ? lines[1].substring(0, 80)  : "";
  }

  const statMatch = text.match(/\$[\d,.]+[BMK]?/);
  if (statMatch && result.type === "stat") {
    result.mainText = statMatch[0];
    const surrounding = text.substring(
      Math.max(0, text.indexOf(statMatch[0]) + statMatch[0].length),
      text.indexOf(statMatch[0]) + statMatch[0].length + 60
    ).trim().split(".")[0];
    result.subText = surrounding || result.subText;
  }

  return result;
}


// ══════════════════════════════════════════════════════════════════════════════
// REMOTION_DATA VALIDATION
// Checks all render payloads for missing/invalid data before hitting the server.
// Returns array of { sceneNum, type, issue } objects.
// ══════════════════════════════════════════════════════════════════════════════
function validateRemotionPayloads_(scenes, payloads) {
  const issues = [];

  scenes.forEach((scene, idx) => {
    const p = payloads[idx];
    if (!p) return; // null payload = scene was skipped intentionally (missing date etc.)

    const props = p.props || {};
    const num   = scene.sceneNum;
    const type  = scene.sceneType;

    switch (p.compositionId.replace(/-Vertical$/, "")) {

      case "CheckpointCard":
        if (!props.date  || props.date.toString().trim()  === "") issues.push({ sceneNum: num, type, issue: "Missing checkpoint date" });
        if (!props.event || props.event.toString().trim() === "") issues.push({ sceneNum: num, type, issue: "Missing checkpoint event" });
        break;

      case "TextImpactScene":
        if (!props.mainText || props.mainText.toString().trim() === "") issues.push({ sceneNum: num, type, issue: "mainText is empty — scene will render blank" });
        break;

      case "InfographicScene": {
        const validTypes = ["data_callout","counter_animation","line_graph","bar_chart","split_comparison","before_after_card"];
        if (!props.type || !validTypes.includes(props.type.toLowerCase())) {
          issues.push({ sceneNum: num, type, issue: "Unknown infographic type: '" + (props.type || "none") + "'" });
        }
        if ((props.type === "bar_chart" || props.type === "line_graph") &&
            (!props.dataPoints || props.dataPoints.length === 0)) {
          issues.push({ sceneNum: num, type, issue: props.type + " has no data points" });
        }
        if (props.type === "data_callout" && (!props.value || props.value.toString().trim() === "")) {
          issues.push({ sceneNum: num, type, issue: "data_callout missing value" });
        }
        break;
      }

      case "TimelineReveal":
        if (!props.checkpoints || props.checkpoints.length < 2) {
          issues.push({ sceneNum: num, type, issue: "Timeline needs at least 2 checkpoints (found " + (props.checkpoints ? props.checkpoints.length : 0) + ")" });
        }
        break;

      case "RiskMatrix":
        if (!props.risks || props.risks.length === 0) {
          issues.push({ sceneNum: num, type, issue: "Risk Matrix has no risks — check REMOTION_DATA in Visual Library col 19" });
        } else {
          const bad = props.risks.filter(r => r.likelihood < 1 || r.likelihood > 3 || r.impact < 1 || r.impact > 3);
          if (bad.length > 0) issues.push({ sceneNum: num, type, issue: bad.length + " risk(s) have likelihood/impact outside 1–3 range" });
        }
        break;

      case "KPIDashboard":
        if (!props.kpis || props.kpis.length === 0) {
          issues.push({ sceneNum: num, type, issue: "KPI Dashboard has no KPIs — check REMOTION_DATA in Visual Library col 19" });
        } else {
          const bad = props.kpis.filter(k => !k.label || !k.value);
          if (bad.length > 0) issues.push({ sceneNum: num, type, issue: bad.length + " KPI(s) missing label or value" });
        }
        break;

      case "ProgressGauge":
        if (!props.gauges || props.gauges.length === 0) {
          issues.push({ sceneNum: num, type, issue: "Progress Gauge has no gauges — check REMOTION_DATA in Visual Library col 19" });
        } else {
          const bad = props.gauges.filter(g => !g.label || g.value < 0 || g.value > 100 || isNaN(g.value));
          if (bad.length > 0) issues.push({ sceneNum: num, type, issue: bad.length + " gauge(s) missing label or value not in 0–100 range" });
        }
        break;
    }
  });

  return issues;
}


// ══════════════════════════════════════════════════════════════════════════════
// CHECK RENDERER HEALTH
// ══════════════════════════════════════════════════════════════════════════════
function checkRendererHealth(ui) {
  try {
    const response = UrlFetchApp.fetch(getRemotionServerUrl() + "/health", {
      muteHttpExceptions: true,
      followRedirects   : true
    });
    if (response.getResponseCode() === 200) {
      Logger.log("Stage 8D: Remotion server health check passed");
      return true;
    }
  } catch (e) {}

  ui.alert(
    "⚠️ Remotion Server Not Running",
    "The GovernX Remotion renderer is not running.\n\n" +
    "To start it:\n" +
    "1. Open Terminal on your computer\n" +
    "2. Navigate to the governx-remotion folder:\n" +
    "   cd path/to/governx-remotion\n" +
    "3. Run: npm start\n" +
    "4. Wait for 'GovernX Remotion Renderer — RUNNING'\n" +
    "5. Come back and run Stage 8D again.\n\n" +
    "Server address: " + getRemotionServerUrl(),
    ui.ButtonSet.OK
  );
  return false;
}


// ══════════════════════════════════════════════════════════════════════════════
// RENDER SINGLE SCENE
// Shows all available scenes with status, validates data, renders chosen scene.
// Re-prompts after success so user can queue multiple single-scene renders.
// ══════════════════════════════════════════════════════════════════════════════
function renderSingleRemotionScene() {

  const idea = getActiveIdeaRow();
  if (!idea) return;

  const ui = SpreadsheetApp.getUi();
  if (!checkRendererHealth(ui)) return;

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const visualSheet = ss.getSheetByName(SHEET.VISUAL);
  const visualData  = visualSheet.getDataRange().getValues();
  const master      = getMasterContent(idea.id);

  // ── Build full scene list for this content ID ──────────────────────────────
  let checkpointTotal = 0;
  for (let i = 1; i < visualData.length; i++) {
    if (visualData[i][COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    if (visualData[i][COL_VISUAL.SCENE_TYPE - 1] === "Checkpoint") checkpointTotal++;
  }

  const allScenes = [];
  let checkpointNum = 0;
  for (let i = 1; i < visualData.length; i++) {
    const row = visualData[i];
    if (row[COL_VISUAL.ID - 1].toString().trim() !== idea.id) continue;
    if (row[COL_VISUAL.SCENE_TYPE - 1] === "Checkpoint") checkpointNum++;
    const sceneType = row[COL_VISUAL.SCENE_TYPE - 1].toString().trim();
    if (!REMOTION_SCENE_TYPES.some(t => sceneType.toLowerCase().includes(t.toLowerCase()))) continue;
    allScenes.push({
      row             : i + 1,
      sceneNum        : row[COL_VISUAL.SCENE_NUM        - 1].toString().trim(),
      sceneType,
      status          : row[COL_VISUAL.STATUS           - 1].toString().trim(),
      description     : row[COL_VISUAL.DESCRIPTION      - 1].toString(),
      assemblyNotes   : row[COL_VISUAL.ASSEMBLY_NOTES   - 1].toString(),
      checkpointDate  : row[COL_VISUAL.CHECKPOINT_DATE  - 1].toString().trim(),
      checkpointEvent : row[COL_VISUAL.CHECKPOINT_EVENT - 1].toString().trim(),
      checkpointAngle : row[COL_VISUAL.CHECKPOINT_ANGLE - 1].toString().trim(),
      checkpointNum,
      checkpointTotal,
      remotionData    : row.length > 18 ? row[COL_VISUAL_EXTENDED.REMOTION_DATA  - 1].toString().trim() : "",
      remotionStyle   : row.length > 19 ? row[COL_VISUAL_EXTENDED.REMOTION_STYLE - 1].toString().trim() : "",
      voiceoverSync   : row.length > 20 ? row[COL_VISUAL_EXTENDED.VOICEOVER_SYNC - 1].toString().trim() : "",
      sceneScore      : row.length > 21 ? row[COL_VISUAL_EXTENDED.SCENE_SCORE    - 1].toString().trim() : ""
    });
  }

  if (allScenes.length === 0) {
    ui.alert("No Remotion scenes found for: " + idea.id + "\nRun Stages 4 and 4B first.");
    return;
  }

  // ── Loop: render scenes until user cancels ─────────────────────────────────
  while (true) {

    // Build scene list text for the prompt
    const sceneList = allScenes.map(s => {
      const statusIcon = s.status === "Done" ? "✅" : s.status === "Error" ? "❌" : s.status === "Skip" ? "⏭" : "⬜";
      const hasData    = s.remotionData && s.remotionData.length > 5 ? "📊" : "  ";
      return statusIcon + " " + hasData + " Scene " + s.sceneNum.toString().padEnd(4) + " | " + s.sceneType;
    }).join("\n");

    const sceneResponse = ui.prompt(
      "🎬 Render Single Scene — " + idea.company,
      "Available scenes:\n\n" + sceneList + "\n\n" +
      "✅ Done  ❌ Error  ⏭ Skip  ⬜ Pending  📊 Has REMOTION_DATA\n\n" +
      "Enter scene number to render (e.g. 1, 2, I1):\n" +
      "(Cancel to exit)",
      ui.ButtonSet.OK_CANCEL
    );

    if (sceneResponse.getSelectedButton() !== ui.Button.OK) return;

    const targetSceneNum = sceneResponse.getResponseText().trim();
    if (!targetSceneNum) continue;

    const targetScene = allScenes.find(s => s.sceneNum === targetSceneNum);
    if (!targetScene) {
      ui.alert("Scene '" + targetSceneNum + "' not found. Check the list and try again.");
      continue;
    }

    // ── Validate before sending ──────────────────────────────────────────────
    const payload = buildRenderPayload(targetScene, idea, master, allScenes);
    if (!payload) {
      ui.alert(
        "⚠️ Cannot Build Payload",
        "Scene " + targetSceneNum + " (" + targetScene.sceneType + ") could not build a render payload.\n\n" +
        "Common causes:\n" +
        "• Checkpoint scene missing date or event\n" +
        "• REMOTION_DATA column (col 19) is empty\n\n" +
        "Fix the data in Visual Library and try again.",
        ui.ButtonSet.OK
      );
      continue;
    }

    const issues = validateRemotionPayloads_([targetScene], [payload]);
    if (issues.length > 0) {
      const proceed = ui.alert(
        "⚠️ Data Validation Issue",
        "Scene " + targetSceneNum + " has a data problem:\n\n" +
        issues.map(v => "• " + v.issue).join("\n") + "\n\n" +
        "Render anyway?",
        ui.ButtonSet.YES_NO
      );
      if (proceed !== ui.Button.YES) continue;
    }

    // ── Render ───────────────────────────────────────────────────────────────
    try {
      visualSheet.getRange(targetScene.row, COL_VISUAL.STATUS).setValue("Submitted");
      SpreadsheetApp.flush();

      const response = UrlFetchApp.fetch(getRemotionServerUrl() + "/render", {
        method            : "post",
        contentType       : "application/json",
        payload           : JSON.stringify(payload),
        muteHttpExceptions: true
      });

      const body = JSON.parse(response.getContentText());

      if (body.success) {
        const scenesFolder = getOrCreateScenesFolder(idea.id, idea.company);
        const serverBase   = getRemotionServerUrl();
        const mp4Url       = body.url.replace(/https?:\/\/localhost:\d+/, serverBase)
                                      .replace(/https?:\/\/127\.0\.0\.1:\d+/, serverBase);
        const mp4Response  = UrlFetchApp.fetch(mp4Url, {
          muteHttpExceptions: true,
          headers: { "ngrok-skip-browser-warning": "true" }
        });
        const blob      = mp4Response.getBlob().setContentType("video/mp4").setName(body.filename || "scene.mp4");
        const driveFile = scenesFolder.createFile(blob);

        visualSheet.getRange(targetScene.row, COL_VISUAL.STATUS        ).setValue("Done");
        visualSheet.getRange(targetScene.row, COL_VISUAL.AI_CLIP_URL   ).setValue(driveFile.getUrl());
        visualSheet.getRange(targetScene.row, COL_VISUAL.BUILT_WHERE   ).setValue("Remotion — Local");
        SpreadsheetApp.flush();

        // Update in-memory status so the scene list reflects it next loop
        targetScene.status = "Done";

        ui.alert(
          "✅ Scene " + targetSceneNum + " Rendered",
          "File: " + body.filename + "\n" +
          "Drive: " + driveFile.getUrl() + "\n\n" +
          "Render another scene? Click OK on the next screen to continue,\nor Cancel to exit.",
          ui.ButtonSet.OK
        );

      } else {
        visualSheet.getRange(targetScene.row, COL_VISUAL.STATUS).setValue("Error");
        SpreadsheetApp.flush();
        targetScene.status = "Error";
        throw new Error(body.error || "Unknown render error");
      }

    } catch (err) {
      ui.alert(
        "❌ Scene " + targetSceneNum + " Failed",
        err.message + "\n\nYou can fix the data and try again from the scene list.",
        ui.ButtonSet.OK
      );
    }

  } // end while loop
}


// ══════════════════════════════════════════════════════════════════════════════
// REMOTION DATA PARSER
// Input:  "type=LINE_GRAPH | label=Market Share | unit=% | points=2009:50,2016:1"
// Output: { type: "LINE_GRAPH", label: "Market Share", unit: "%", points: "..." }
// ══════════════════════════════════════════════════════════════════════════════
function parseRemotionData(remotionDataStr) {
  if (!remotionDataStr) return {};
  const result = {};
  const pairs  = remotionDataStr.split("|").map(p => p.trim());
  pairs.forEach(pair => {
    const eqIdx = pair.indexOf("=");
    if (eqIdx === -1) return;
    const key = pair.substring(0, eqIdx).trim().toLowerCase().replace(/[^a-z0-9_]/g, "_");
    const val = pair.substring(eqIdx + 1).trim();
    result[key] = val;
  });
  return result;
}


// ── parseStyleKey ─────────────────────────────────────────────────────────────
function parseStyleKey(remotionStyleStr, key) {
  if (!remotionStyleStr) return null;
  const pairs = remotionStyleStr.split("|").map(p => p.trim());
  for (const pair of pairs) {
    const eqIdx = pair.indexOf("=");
    if (eqIdx === -1) continue;
    const k = pair.substring(0, eqIdx).trim().toLowerCase();
    const v = pair.substring(eqIdx + 1).trim();
    if (k === key.toLowerCase()) return v;
  }
  return null;
}


// ── buildInfographicPropsFromData ─────────────────────────────────────────────
function buildInfographicPropsFromData(rd, variant, scene, idea) {
  const type          = (variant || rd.type || "DATA_CALLOUT").toUpperCase();
  const voiceoverSync = scene.voiceoverSync || "";
  const duration      = parseInt(rd.duration) || 4;

  if (type === "LINE_GRAPH") {
    const points = (rd.points || "").split(",").map(p => {
      const [yr, val] = p.trim().split(":");
      return { year: yr ? yr.trim() : "", value: isNaN(parseFloat(val)) ? 0 : parseFloat(val) };
    }).filter(p => p.year);
    return {
      type      : "line_graph",
      title     : rd.label || idea.company,
      unit      : rd.unit  || "%",
      dataPoints: points,
      highlight : rd.highlight || "",
      voiceoverSync, duration
    };
  }

  if (type === "SPLIT_COMPARISON") {
    const parseVals = (str) => (str || "").split(",").map(p => {
      const ci = p.lastIndexOf(":");
      if (ci === -1) return { label: p.trim(), value: "" };
      return { label: p.substring(0, ci).trim(), value: p.substring(ci + 1).trim() };
    });
    return {
      type        : "split_comparison",
      leftLabel   : rd.left_label  || "Left",
      leftValues  : parseVals(rd.left_values),
      rightLabel  : rd.right_label || "Right",
      rightValues : parseVals(rd.right_values),
      bottomNote  : rd.bottom_note || "",
      voiceoverSync, duration
    };
  }

  if (type === "DATA_CALLOUT") {
    return {
      type    : "data_callout",
      value   : rd.value   || "",
      label   : rd.label   || "",
      context : rd.context || "",
      voiceoverSync, duration
    };
  }

  if (type === "COUNTER_ANIMATION") {
    return {
      type  : "counter_animation",
      from  : rd.from  || "0",
      to    : rd.to    || "0",
      unit  : rd.unit  || "",
      label : rd.label || "",
      voiceoverSync, duration
    };
  }

  if (type === "BEFORE_AFTER_CARD") {
    // Row separator is ";;" — NOT "|" (pipe is reserved as the REMOTION_DATA field delimiter)
    // Format: before_rows=Audit scope→75%;;IT systems→Not covered;;Risk appetite→Not defined
    const parseRows = (str) => (str || "").split(";;").map(r => {
      const parts = r.split("→").map(p => p.trim());
      return { item: parts[0] || "", value: parts[1] || "" };
    }).filter(r => r.item);
    return {
      type        : "before_after_card",
      beforeLabel : rd.before_label || "Standard",
      beforeRows  : parseRows(rd.before_rows),
      afterLabel  : rd.after_label  || "Actual",
      afterRows   : parseRows(rd.after_rows),
      verdict     : rd.verdict      || "",
      voiceoverSync, duration
    };
  }

  if (type === "BAR_CHART") {
    // REMOTION_DATA: type=BAR_CHART | title=My Title | unit=$B | points=Label:value:highlight,...
    const points = (rd.points || "").split(",").map(p => {
      const parts = p.trim().split(":");
      return {
        label    : parts[0] ? parts[0].trim() : "",
        value    : isNaN(parseFloat(parts[1])) ? 0 : parseFloat(parts[1]),
        highlight: (parts[2] || "").trim().toLowerCase() === "true"
      };
    }).filter(p => p.label);
    return {
      type      : "bar_chart",
      title     : rd.title || rd.label || "",
      unit      : rd.unit  || "",
      dataPoints: points.length > 0 ? points : [],
      voiceoverSync, duration
    };
  }

  Logger.log("Stage 8D: Unknown infographic variant '" + type + "' — using data_callout fallback");
  return {
    type    : "data_callout",
    value   : rd.value || rd.to || "",
    label   : rd.label || idea.company,
    context : "",
    voiceoverSync, duration
  };
}


// ══════════════════════════════════════════════════════════════════════════════
// CALC SCENE DURATION FROM VOICEOVER WORD COUNT
// Rate: 138 wpm (ElevenLabs eleven_multilingual_v2 at default settings)
// Buffer: +0.8s tail so the visual doesn't cut before the last word lands
// Clamp: min 3s, max 15s — then convert to frames at 30fps
// Falls back to defaultSeconds when VOICEOVER_SYNC is empty.
// ══════════════════════════════════════════════════════════════════════════════
function calcSceneDuration(voiceoverText, defaultSeconds) {
  const text = (voiceoverText || "").trim();
  if (!text) return Math.round((defaultSeconds || 5) * 30);

  const wordCount      = text.split(/\s+/).filter(w => w.length > 0).length;
  const rawSeconds     = (wordCount / 2.3) + 0.8;   // 138 wpm + 0.8s tail
  const clampedSeconds = Math.min(15, Math.max(3, rawSeconds));
  return Math.round(clampedSeconds * 30);            // frames at 30fps
}


function cleanCheckpointDate(rawDate) {
  if (!rawDate) return "";
  const s = rawDate.toString().trim();
  if (s.length <= 20 && !s.includes("GMT")) return s;
  const yearMatch = s.match(/\b(19|20)\d{2}\b/);
  if (yearMatch) return yearMatch[0];
  return s;
}