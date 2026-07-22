/* ============================================================================
   dashboard.gs — GovernX Content OS
   Creates and populates the Dashboard tab with live COUNTIF/QUERY formulas
   Run once to build — formulas auto-update as data changes
   ============================================================================ */

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // ── Get or create Dashboard tab ──────────────────────────────────────────
  let dash = ss.getSheetByName("Dashboard");
  if (dash) {
    const response = ui.alert(
      "Dashboard Already Exists",
      "Rebuild the Dashboard tab? All existing content will be replaced.",
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    dash.clear();
    dash.clearFormats();
  } else {
    dash = ss.insertSheet("Dashboard");
    // Move to first position
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);
  }

  // ── Color palette — clean light professional theme ────────────────────────
  const C = {
    BLACK      : "#FFFFFF",   // sheet background → white
    RED        : "#CC0000",   // GovernX red (slightly deeper for light bg)
    WHITE      : "#111111",   // primary text → near black
    DARK_GRAY  : "#F8F9FA",   // metric card bg → light gray
    MID_GRAY   : "#F1F3F5",   // alternate row
    LIGHT_GRAY : "#F5F5F5",
    BORDER     : "#DEE2E6",   // subtle border color
    GREEN      : "#16A34A",   // green for done/ready
    AMBER      : "#D97706",   // amber for pending
    BLUE       : "#2563EB",   // blue for info
    MUTED      : "#6B7280",   // muted text
    HEADER_BG  : "#1E293B",   // section header → dark navy (contrast)
    ROW_ALT    : "#F1F3F5",   // alternate row bg
    DONE_BG    : "#DCFCE7",   // done badge bg
    DONE_FG    : "#15803D",   // done badge text
    DRAFT_BG   : "#FEF9C3",   // draft badge bg
    DRAFT_FG   : "#A16207",   // draft badge text
    ERROR_BG   : "#FEE2E2",   // error row bg
    ERROR_FG   : "#DC2626"    // error text
  };

  // ── Sheet setup ───────────────────────────────────────────────────────────
  dash.setTabColor(C.RED);
  dash.setFrozenRows(0);

  // Hide gridlines
  dash.setHiddenGridlines(true);

  // Set column widths
  const colWidths = [20, 180, 140, 100, 80, 80, 80, 80, 80, 80, 80, 80, 80, 80, 140, 20];
  colWidths.forEach((w, i) => dash.setColumnWidth(i + 1, w));
  dash.setRowHeight(1, 20);

  let row = 2;

  // ── Helper: write section header ──────────────────────────────────────────
  const sectionHeader = (text, startRow) => {
    dash.getRange(startRow, 2, 1, 14).merge()
      .setValue(text.toUpperCase())
      .setFontSize(9)
      .setFontWeight("bold")
      .setFontColor(C.RED)
      .setBackground(C.BLACK)
      .setFontFamily("Montserrat");
    dash.setRowHeight(startRow, 28);
    return startRow + 1;
  };

  // ── Helper: styled cell ───────────────────────────────────────────────────
  const cell = (r, c) => dash.getRange(r, c);
  const range = (r, c, rows, cols) => dash.getRange(r, c, rows, cols);

  // ════════════════════════════════════════════════════════════════════════════
  // HEADER BANNER
  // ════════════════════════════════════════════════════════════════════════════
  range(row, 2, 2, 14).merge()
    .setValue("🎬  GovernX Content OS — Dashboard")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#1E293B")
    .setFontFamily("Montserrat")
    .setVerticalAlignment("middle");
  dash.setRowHeight(row, 50);
  dash.setRowHeight(row + 1, 10);
  row += 3;

  // Last updated row
  range(row, 2, 1, 14).merge()
    .setFormula('="Last updated: "&TEXT(NOW(),"DD MMM YYYY, HH:MM")')
    .setFontSize(9)
    .setFontColor("#94A3B8")
    .setBackground("#1E293B")
    .setFontFamily("Montserrat");
  dash.setRowHeight(row, 20);
  row += 2;

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 1 — CONTENT OVERVIEW METRICS
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Content Overview", row);

  // Metric cards — row of 6
  const metrics = [
    { label: "Total Ideas",        formula: '=COUNTA(\'Idea Catalogue\'!A2:A1000)' },
    { label: "Scripts Written",    formula: '=COUNTA(\'Script Bank\'!A2:A1000)' },
    { label: "Scenes Generated",   formula: '=COUNTA(\'Visual Library\'!A2:A1000)' },
    { label: "Sources Researched", formula: '=COUNTA(\'Research Database\'!A2:A1000)' },
    // Uploaded = has a YouTube video id (col R); Published = went public (PUBLISH_DATE, col B)
    { label: "Uploaded",           formula: '=COUNTIF(\'Publishing Tracker\'!R2:R1000,"<>")' },
    { label: "Published (public)", formula: '=COUNTIF(\'Publishing Tracker\'!B2:B1000,"<>")' }
  ];

  // Metric label row
  metrics.forEach((m, i) => {
    const col = 2 + (i * 2);
    range(row, col, 1, 2).merge()
      .setValue(m.label)
      .setFontSize(9)
      .setFontColor(C.MUTED)
      .setBackground(C.DARK_GRAY)
      .setFontFamily("Montserrat")
      .setHorizontalAlignment("center");
    dash.setRowHeight(row, 22);
  });
  row++;

  // Metric value row
  metrics.forEach((m, i) => {
    const col = 2 + (i * 2);
    range(row, col, 1, 2).merge()
      .setFormula(m.formula)
      .setFontSize(22)
      .setFontWeight("bold")
      .setFontColor(C.WHITE)
      .setBackground(C.DARK_GRAY)
      .setFontFamily("Montserrat")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    dash.setRowHeight(row, 44);
  });
  row += 2;

  // ── Economics headline (portfolio totals from the Video Economics tab) ──────
  const econCards = [
    { label: "Total Cost",    formula: '=IFERROR(SUM(\'Video Economics\'!H13:H1000),0)', fmt: "$#,##0.00" },
    { label: "Total Revenue", formula: '=IFERROR(SUM(\'Video Economics\'!J13:J1000),0)', fmt: "$#,##0.00" },
    { label: "Net",           formula: '=IFERROR(SUM(\'Video Economics\'!J13:J1000)-SUM(\'Video Economics\'!H13:H1000),0)', fmt: "$#,##0.00" },
    { label: "Portfolio ROI", formula: '=IFERROR((SUM(\'Video Economics\'!J13:J1000)-SUM(\'Video Economics\'!H13:H1000))/SUM(\'Video Economics\'!H13:H1000),"—")', fmt: "0%" }
  ];
  econCards.forEach(function (m, i) {
    const col = 2 + (i * 2);
    range(row, col, 1, 2).merge().setValue(m.label)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(C.DARK_GRAY)
      .setFontFamily("Montserrat").setHorizontalAlignment("center");
    dash.setRowHeight(row, 22);
  });
  row++;
  econCards.forEach(function (m, i) {
    const col = 2 + (i * 2);
    range(row, col, 1, 2).merge().setFormula(m.formula).setNumberFormat(m.fmt)
      .setFontSize(20).setFontWeight("bold").setFontColor(C.WHITE).setBackground(C.DARK_GRAY)
      .setFontFamily("Montserrat").setHorizontalAlignment("center").setVerticalAlignment("middle");
    dash.setRowHeight(row, 44);
  });
  row += 2;

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 2 — PIPELINE STAGE TRACKER
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Pipeline Stage Progress", row);

  // Column headers
  // "S8" used to mean Stage 8D (per-scene MP4s) and lit from Visual Library
  // STATUS="Ready". 8D is retired — the film is assembled in one pass — so that
  // column could never light again. It is now S7B (per-scene narration audio),
  // which is what the assembly actually depends on.
  const stageHeaders = ["ID", "Company", "Domain", "Discipline", "S1", "S2", "S3", "S4", "S5", "S6", "S7", "S7B", "S9", "S10", "Thumb", "S11", "Status"];
  stageHeaders.forEach((h, i) => {
    cell(row, 2 + i)
      .setValue(h)
      .setFontSize(9)
      .setFontWeight("bold")
      .setFontColor(C.MUTED)
      .setBackground(C.HEADER_BG)
      .setFontFamily("Montserrat")
      .setHorizontalAlignment(i >= 4 ? "center" : "left");
  });
  dash.setRowHeight(row, 24);
  row++;

  // Pipeline data rows — formula-driven
  // Each row reads from Master Content Table and checks other sheets
  const pipelineFormulas = [
    // Row formula pattern for each content entry
    // S1: has entry in Master Content Table
    // S2: has entry in Research Database
    // S3: has entry in Script Bank
    // S4: has entry in Visual Library
    // S5: has entry in Publishing Tracker
    // S6: has Production Package link
    // S7: has Voiceover Audio link
    // S8: has any AI Clip URL in Visual Library
    // S9: has Assembly Guide link
  ];

  // Size the table to the real number of ideas (+ headroom) so it never truncates
  // at 10 and new videos appear. Re-run Build Dashboard to relayout after big adds.
  const ideaSheet = ss.getSheetByName("Idea Catalogue");
  const ideaCount = ideaSheet
    ? ideaSheet.getRange("A2:A1000").getValues().filter(function (r) { return String(r[0]).trim() !== ""; }).length
    : 10;
  const N = Math.max(12, ideaCount + 6);
  const firstDataRow = row;

  // Write N pipeline rows with IFERROR/VLOOKUP formulas
  for (let r = 0; r < N; r++) {
    const dataRow = r + 2; // data starts at row 2 in source sheets
    const rowNum = row + r;
    const bg = r % 2 === 0 ? C.DARK_GRAY : C.ROW_ALT;

    // ID
    cell(rowNum, 2)
      .setFormula(`=IFERROR(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),"")`)
      .setFontSize(10).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    // Company
    cell(rowNum, 3)
      .setFormula(`=IFERROR(INDEX('Idea Catalogue'!B$2:B$1000,${dataRow}-1),"")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat").setFontWeight("bold");

    // Domain
    cell(rowNum, 4)
      .setFormula(`=IFERROR(INDEX('Idea Catalogue'!C$2:C$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    // Discipline (from Master Content Table col 11)
    cell(rowNum, 5)
      .setFormula(`=IFERROR(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Master Content Table'!A:K,11,FALSE),"")`)
      .setFontSize(9).setFontColor(C.RED).setBackground(bg).setFontFamily("Montserrat").setFontWeight("bold").setHorizontalAlignment("center");

    // Stage indicators using COUNTIF to check presence of ID in each sheet
    const stages = [
      `=IFERROR(IF(COUNTIF('Master Content Table'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"")`,
      `=IFERROR(IF(COUNTIF('Research Database'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"")`,
      `=IFERROR(IF(COUNTIF('Script Bank'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"")`,
      `=IFERROR(IF(COUNTIF('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"")`,
      `=IFERROR(IF(COUNTIF('Publishing Tracker'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"")`,
      `=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:M,13,FALSE)<>"","✓","·"),"·")`,
      `=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:N,14,FALSE)<>"","✓","·"),"·")`,
      // S7B — per-scene narration audio (Visual Library col W = VOICEOVER_AUDIO_URL).
      // This is what Assemble Film downloads; without it every scene renders silent.
      `=IFERROR(IF(COUNTIF('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,IF(COUNTA(FILTER('Visual Library'!W:W,'Visual Library'!A:A=INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1)))>0,"✓","~"),"·"),"·")`,
      // S9 — the assembled film. Stage 9C writes its Drive link to Publishing
      // Tracker col O (15 = SCENES_FOLDER); this used to read col 16 (Assembly
      // Guide), which Shotstack filled, so it never lit for a Remotion assembly.
      `=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:O,15,FALSE)<>"","✓","·"),"·")`,
      // S10 — YouTube metadata generated (a row exists in the YouTube Metadata tab)
      `=IFERROR(IF(COUNTIF('YouTube Metadata'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1))>0,"✓","·"),"·")`,
      // Thumb — a thumbnail link was written to Publishing Tracker col D
      `=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:D,4,FALSE)<>"","✓","·"),"·")`,
      // S11 — uploaded to YouTube (video id in Publishing Tracker col R)
      `=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:R,18,FALSE)<>"","✓","·"),"·")`
    ];

    stages.forEach((formula, si) => {
      cell(rowNum, 6 + si)
        .setFormula(formula)
        .setFontSize(12)
        .setFontWeight("bold")
        .setBackground(bg)
        .setFontFamily("Montserrat")
        .setHorizontalAlignment("center");
    });

    // Status — reflects the real end states (assembled → uploaded → published)
    const idExpr = `INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1)`;
    cell(rowNum, 6 + stages.length)
      .setFormula(
        `=IFERROR(IF(${idExpr}="","",` +
        `IF(VLOOKUP(${idExpr},'Publishing Tracker'!A:B,2,FALSE)<>"","🚀 Published",` +
        `IF(VLOOKUP(${idExpr},'Publishing Tracker'!A:R,18,FALSE)<>"","⬆ Uploaded",` +
        `IF(VLOOKUP(${idExpr},'Publishing Tracker'!A:O,15,FALSE)<>"","🎬 Assembled",` +
        `IF(COUNTIF('Visual Library'!A:A,${idExpr})>0,"🎨 Scened",` +
        `IF(COUNTIF('Script Bank'!A:A,${idExpr})>0,"📝 Scripted",` +
        `"💡 Idea")))))),"")`)
      .setFontSize(10)
      .setFontColor(C.WHITE)
      .setBackground(bg)
      .setFontFamily("Montserrat");

    dash.setRowHeight(rowNum, 28);
  }
  row += N + 2;

  // Add conditional formatting for stage dots — N rows × 12 stage columns (6–17)
  const stageRange = dash.getRange(firstDataRow, 6, N, 12);
  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("✓")
    .setFontColor(C.GREEN)
    .setRanges([stageRange])
    .build();
  const amberRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("~")
    .setFontColor(C.AMBER)
    .setRanges([stageRange])
    .build();
  const grayRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("·")
    .setFontColor(C.BORDER)
    .setRanges([stageRange])
    .build();
  dash.setConditionalFormatRules([greenRule, amberRule, grayRule]);

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 3 — CONTENT BREAKDOWN
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Content Library Breakdown", row);

  // Discipline breakdown
  range(row, 2, 1, 4).merge()
    .setValue("By Discipline")
    .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
    .setBackground(C.HEADER_BG).setFontFamily("Montserrat");

  range(row, 7, 1, 4).merge()
    .setValue("By Domain")
    .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
    .setBackground(C.HEADER_BG).setFontFamily("Montserrat");

  range(row, 12, 1, 4).merge()
    .setValue("By Type")
    .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
    .setBackground(C.HEADER_BG).setFontFamily("Montserrat");
  dash.setRowHeight(row, 24);
  row++;

  const breakdowns = [
    { values: ["GRC", "BPR", "GRC+BPR"], formula: (v) => `=COUNTIF('Master Content Table'!K:K,"${v}")`, col: 2 },
    { values: ["Business", "Sports", "Media & Creator Economy", "Public Sector", "Startups & Tech"], formula: (v) => `=COUNTIF('Master Content Table'!C:C,"${v}")`, col: 7 },
    { values: ["Success", "Failure", "Turning Point", "Collapse", "Underdog"], formula: (v) => `=COUNTIF('Master Content Table'!F:F,"${v}")`, col: 12 }
  ];

  const maxBreakdownRows = 5;
  for (let r = 0; r < maxBreakdownRows; r++) {
    const rowNum = row + r;
    const bg = r % 2 === 0 ? C.DARK_GRAY : C.ROW_ALT;
    dash.setRowHeight(rowNum, 24);

    breakdowns.forEach(bd => {
      if (r < bd.values.length) {
        range(rowNum, bd.col, 1, 3).merge()
          .setValue(bd.values[r])
          .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat");
        cell(rowNum, bd.col + 3)
          .setFormula(bd.formula(bd.values[r]))
          .setFontSize(10).setFontWeight("bold").setFontColor(C.RED)
          .setBackground(bg).setFontFamily("Montserrat").setHorizontalAlignment("center");
      } else {
        range(rowNum, bd.col, 1, 4).merge().setBackground(bg);
      }
    });
  }
  row += maxBreakdownRows + 2;

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 4 — ASSET READINESS
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Asset Readiness", row);

  const assetHeaders = ["ID", "Company", "Total Scenes", "Scenes Ready", "Scenes Pending", "Voiceover", "Package", "Assembly", "Next Action"];
  assetHeaders.forEach((h, i) => {
    range(row, 2 + i, 1, i === 1 ? 2 : 1)
      .setValue(h)
      .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
      .setBackground(C.HEADER_BG).setFontFamily("Montserrat")
      .setHorizontalAlignment(i >= 2 ? "center" : "left");
  });
  dash.setRowHeight(row, 24);
  row++;

  for (let r = 0; r < 10; r++) {
    const dataRow = r + 2;
    const rowNum = row + r;
    const bg = r % 2 === 0 ? C.DARK_GRAY : C.ROW_ALT;
    dash.setRowHeight(rowNum, 26);

    cell(rowNum, 2)
      .setFormula(`=IFERROR(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    range(rowNum, 3, 1, 2).merge()
      .setFormula(`=IFERROR(INDEX('Idea Catalogue'!B$2:B$1000,${dataRow}-1),"")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat").setFontWeight("bold");

    // Total scenes
    cell(rowNum, 5)
      .setFormula(`=IFERROR(COUNTIF('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1)),0)`)
      .setFontSize(10).setFontColor(C.MUTED).setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Scenes ready
    cell(rowNum, 6)
      .setFormula(`=IFERROR(COUNTIFS('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Visual Library'!I:I,"Ready"),0)`)
      .setFontSize(10).setFontColor(C.GREEN).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Scenes pending
    cell(rowNum, 7)
      .setFormula(`=IFERROR(COUNTIFS('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Visual Library'!I:I,"Needed"),0)`)
      .setFontSize(10).setFontColor(C.AMBER).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Voiceover
    cell(rowNum, 8)
      .setFormula(`=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:N,14,FALSE)<>"","✓","—"),"—")`)
      .setFontSize(11).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Package
    cell(rowNum, 9)
      .setFormula(`=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:M,13,FALSE)<>"","✓","—"),"—")`)
      .setFontSize(11).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Assembly
    cell(rowNum, 10)
      .setFormula(`=IFERROR(IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:P,16,FALSE)<>"","✓","—"),"—")`)
      .setFontSize(11).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Next action
    cell(rowNum, 11)
      .setFormula(
        `=IFERROR(IF(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1)="","",` +
        `IF(VLOOKUP(INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Publishing Tracker'!A:P,16,FALSE)="","Run Stage 9 →",` +
        `IF(COUNTIFS('Visual Library'!A:A,INDEX('Idea Catalogue'!A$2:A$1000,${dataRow}-1),'Visual Library'!I:I,"Needed")>0,"Complete scenes →",` +
        `"Edit in CapCut →"))),"")`)
      .setFontSize(9).setFontColor(C.RED).setBackground(bg).setFontFamily("Montserrat").setFontWeight("bold");
  }
  row += 12;

  // Conditional formatting for check marks
  const checkRange = dash.getRange(row - 12, 8, 10, 3);
  const checkGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("✓")
    .setFontColor(C.GREEN)
    .setRanges([checkRange])
    .build();
  const checkGray = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("—")
    .setFontColor(C.BORDER)
    .setRanges([checkRange])
    .build();
  const existingRules = dash.getConditionalFormatRules();
  dash.setConditionalFormatRules([...existingRules, checkGreen, checkGray]);

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 5 — PERFORMANCE TRACKER
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Performance Tracker", row);

  const perfHeaders = ["ID", "Title", "Publish Date", "Views", "CTR %", "CTR vs 6%", "Retention %", "Ret vs 45%", "Decision"];
  perfHeaders.forEach((h, i) => {
    const col = 2 + i;
    const span = i === 1 ? 2 : 1;
    range(row, col, 1, span).merge()
      .setValue(h)
      .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
      .setBackground(C.HEADER_BG).setFontFamily("Montserrat")
      .setHorizontalAlignment(i >= 3 ? "center" : "left");
  });
  dash.setRowHeight(row, 24);
  row++;

  for (let r = 0; r < 10; r++) {
    const dataRow = r + 2;
    const rowNum = row + r;
    const bg = r % 2 === 0 ? C.DARK_GRAY : C.ROW_ALT;
    dash.setRowHeight(rowNum, 26);

    // ID
    cell(rowNum, 2)
      .setFormula(`=IFERROR(INDEX('Publishing Tracker'!A$2:A$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    // Title
    range(rowNum, 3, 1, 2).merge()
      .setFormula(`=IFERROR(INDEX('Publishing Tracker'!C$2:C$1000,${dataRow}-1),"")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat");

    // Publish Date
    cell(rowNum, 5)
      .setFormula(`=IFERROR(INDEX('Publishing Tracker'!B$2:B$1000,${dataRow}-1),"Unpublished")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    // Views
    cell(rowNum, 6)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!E$2:E$1000,${dataRow}-1)="","—",INDEX('Publishing Tracker'!E$2:E$1000,${dataRow}-1)),"—")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // CTR
    cell(rowNum, 7)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!F$2:F$1000,${dataRow}-1)="","—",INDEX('Publishing Tracker'!F$2:F$1000,${dataRow}-1)&"%"),"—")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // CTR benchmark
    cell(rowNum, 8)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!F$2:F$1000,${dataRow}-1)="","—",IF(INDEX('Publishing Tracker'!F$2:F$1000,${dataRow}-1)>=6,"▲ Scale","▼ Watch")),"—")`)
      .setFontSize(9).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Retention
    cell(rowNum, 9)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!G$2:G$1000,${dataRow}-1)="","—",INDEX('Publishing Tracker'!G$2:G$1000,${dataRow}-1)&"%"),"—")`)
      .setFontSize(10).setFontColor(C.WHITE).setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Retention benchmark
    cell(rowNum, 10)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!G$2:G$1000,${dataRow}-1)="","—",IF(INDEX('Publishing Tracker'!G$2:G$1000,${dataRow}-1)>=45,"▲ Scale","▼ Watch")),"—")`)
      .setFontSize(9).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");

    // Decision
    cell(rowNum, 11)
      .setFormula(`=IFERROR(IF(INDEX('Publishing Tracker'!L$2:L$1000,${dataRow}-1)="","Pending",INDEX('Publishing Tracker'!L$2:L$1000,${dataRow}-1)),"Pending")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");
  }
  row += 12;

  // ════════════════════════════════════════════════════════════════════════════
  // SECTION 6 — ERROR LOG SUMMARY
  // ════════════════════════════════════════════════════════════════════════════
  row = sectionHeader("Error Log Summary", row);

  // Error counts
  range(row, 2, 1, 3).merge()
    .setValue("Total errors logged")
    .setFontSize(10).setFontColor(C.MUTED).setBackground(C.DARK_GRAY).setFontFamily("Montserrat");
  cell(row, 5)
    .setFormula(`=IFERROR(COUNTA('Error Log'!A2:A1000)-1,0)`)
    .setFontSize(14).setFontWeight("bold").setFontColor(C.WHITE).setBackground(C.DARK_GRAY).setFontFamily("Montserrat").setHorizontalAlignment("center");

  range(row, 7, 1, 3).merge()
    .setValue("Resolved")
    .setFontSize(10).setFontColor(C.MUTED).setBackground(C.DARK_GRAY).setFontFamily("Montserrat");
  cell(row, 10)
    .setFormula(`=IFERROR(COUNTIF('Error Log'!F2:F1000,TRUE),0)`)
    .setFontSize(14).setFontWeight("bold").setFontColor(C.GREEN).setBackground(C.DARK_GRAY).setFontFamily("Montserrat").setHorizontalAlignment("center");

  range(row, 12, 1, 3).merge()
    .setValue("Unresolved")
    .setFontSize(10).setFontColor(C.MUTED).setBackground(C.DARK_GRAY).setFontFamily("Montserrat");
  cell(row, 15)
    .setFormula(`=IFERROR(COUNTIF('Error Log'!F2:F1000,FALSE),0)`)
    .setFontSize(14).setFontWeight("bold").setFontColor(C.RED).setBackground(C.DARK_GRAY).setFontFamily("Montserrat").setHorizontalAlignment("center");
  dash.setRowHeight(row, 40);
  row += 2;

  // Recent errors table
  const errHeaders = ["Timestamp", "Stage", "Content ID", "Error Type", "Resolved"];
  errHeaders.forEach((h, i) => {
    const colSpan = i === 3 ? 4 : i === 1 ? 3 : 2;
    range(row, 2 + [0,2,5,7,11][i], 1, colSpan).merge()
      .setValue(h)
      .setFontSize(9).setFontWeight("bold").setFontColor(C.MUTED)
      .setBackground(C.HEADER_BG).setFontFamily("Montserrat");
  });
  dash.setRowHeight(row, 24);
  row++;

  for (let r = 0; r < 5; r++) {
    const dataRow = r + 2;
    const rowNum = row + r;
    const bg = r % 2 === 0 ? C.DARK_GRAY : C.ROW_ALT;
    dash.setRowHeight(rowNum, 24);

    range(rowNum, 2, 1, 2).merge()
      .setFormula(`=IFERROR(TEXT(INDEX('Error Log'!A$2:A$1000,${dataRow}-1),"DD MMM"),"")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    range(rowNum, 4, 1, 2).merge()
      .setFormula(`=IFERROR(INDEX('Error Log'!B$2:B$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat");

    range(rowNum, 6, 1, 2).merge()
      .setFormula(`=IFERROR(INDEX('Error Log'!C$2:C$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.MUTED).setBackground(bg).setFontFamily("Montserrat");

    range(rowNum, 8, 1, 4).merge()
      .setFormula(`=IFERROR(INDEX('Error Log'!D$2:D$1000,${dataRow}-1),"")`)
      .setFontSize(9).setFontColor(C.WHITE).setBackground(bg).setFontFamily("Montserrat");

    cell(rowNum, 12)
      .setFormula(`=IFERROR(IF(INDEX('Error Log'!F$2:F$1000,${dataRow}-1)=TRUE,"✓","✗"),"")`)
      .setFontSize(11).setFontWeight("bold").setBackground(bg).setHorizontalAlignment("center").setFontFamily("Montserrat");
  }
  row += 7;

  // ── Footer ────────────────────────────────────────────────────────────────
  range(row, 2, 1, 14).merge()
    .setValue("GovernX Content OS — Reverse-engineering leadership decisions to reveal the urgency of GRC and BPR")
    .setFontSize(9).setFontColor("#94A3B8").setBackground("#1E293B")
    .setFontFamily("Montserrat").setHorizontalAlignment("center");
  dash.setRowHeight(row, 30);

  // ── Set background for entire sheet → white ───────────────────────────────
  dash.getRange(1, 1, row + 5, 16).setBackground("#FFFFFF");

  // ── Freeze first column (spacer) ──────────────────────────────────────────
  dash.setFrozenColumns(1);
  dash.setFrozenRows(0);

  // ── Print settings ────────────────────────────────────────────────────────
  dash.setColumnWidth(1, 20);
  dash.setColumnWidth(16, 20);

  SpreadsheetApp.flush();

  ui.alert(
    "✅ Dashboard Built",
    "The Dashboard tab has been created with live formulas.\n\n" +
    "It auto-updates as you run pipeline stages — no manual refresh needed.\n\n" +
    "To rebuild at any time: GovernX menu → Build Dashboard.",
    ui.ButtonSet.OK
  );
}