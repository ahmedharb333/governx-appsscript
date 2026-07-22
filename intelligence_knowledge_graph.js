/* ============================================================================
   Intelligence_1.2_KnowledgeGraph.gs — GovernX Intelligence Platform
   PHASE 1 · UNIT 1.2 — Knowledge Graph

   Turns the flat company data into a network of typed nodes + weighted edges.

   Tabs created:
     • Knowledge_Graph (nodes — N-<type>-####)
     • Graph_Edges     (edges — Source_Node → Target_Node, weighted)

   Node types: Company, Person, Board, Law, Regulator, Industry, Country,
               Framework, Video, Report.

   SAFETY:
   - Additive. Own constants (SHEET_KG, COL_KG_*). No config.gs edits.
   - Uses intelSS_() (Intelligence_Core) for the spreadsheet handle.
   - Reuses helpers from Intelligence_1.1 (findCompanyRow_, getSelectedCompanyId_,
     normalizeName_, findProfileRow_, INTEL_SYSTEM_CONTEXT,
     callClaudeWithCustomSystem, logError). Keep those files in the project.

   HOW TO USE:
   1. Run  setupKnowledgeGraphTabs()  → creates the 2 tabs.
   2. Run  syncCompanyNodes()          → one Company node per Company_Master row.
   3. Select a company, run  generateGraphFromCompany()
      → Claude emits its people/regulators/industry/failure nodes + edges.
   ============================================================================ */


// ── Tab names ────────────────────────────────────────────────────────────────
const SHEET_KG = {
  NODES : "Knowledge_Graph",
  EDGES : "Graph_Edges"
};

// ── Column maps (1-based) ─────────────────────────────────────────────────────
const COL_KG_NODE = {
  NODE_ID     : 1,
  NODE_TYPE   : 2,
  NODE_NAME   : 3,
  DESCRIPTION : 4,
  REF_ID      : 5   // FK to a source row (e.g. Company_ID) when applicable
};

const COL_KG_EDGE = {
  SOURCE_NODE  : 1,
  RELATIONSHIP : 2,
  TARGET_NODE  : 3,
  WEIGHT       : 4
};

// ── Node type → ID code ───────────────────────────────────────────────────────
const KG_TYPE_CODE = {
  "Company"   : "CO", "Person"    : "PE", "Board"     : "BD",
  "Law"       : "LW", "Regulator" : "RG", "Industry"  : "IN",
  "Country"   : "CN", "Framework" : "FR", "Video"     : "VD",
  "Report"    : "RP"
};
const KG_NODE_TYPES = Object.keys(KG_TYPE_CODE);

const INTEL_HEADER_BG_KG = "#1a1a2e";
const INTEL_HEADER_FG_KG = "#ffffff";


// ══════════════════════════════════════════════════════════════════════════════
// SETUP
// ══════════════════════════════════════════════════════════════════════════════
function setupKnowledgeGraphTabs() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  let nodes = ss.getSheetByName(SHEET_KG.NODES);
  const nodesNew = !nodes;
  if (nodesNew) nodes = ss.insertSheet(SHEET_KG.NODES);
  nodes.getRange(1, 1, 1, 5).setValues([["Node_ID", "Node_Type", "Node_Name", "Description", "Ref_ID"]])
       .setBackground(INTEL_HEADER_BG_KG).setFontColor(INTEL_HEADER_FG_KG).setFontWeight("bold");
  nodes.setFrozenRows(1);
  [140, 120, 240, 420, 140].forEach((w, i) => nodes.setColumnWidth(i + 1, w));
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(KG_NODE_TYPES, true).setAllowInvalid(true).build();
  nodes.getRange(2, COL_KG_NODE.NODE_TYPE, 999, 1).setDataValidation(typeRule);

  let edges = ss.getSheetByName(SHEET_KG.EDGES);
  const edgesNew = !edges;
  if (edgesNew) edges = ss.insertSheet(SHEET_KG.EDGES);
  edges.getRange(1, 1, 1, 4).setValues([["Source_Node", "Relationship", "Target_Node", "Weight"]])
       .setBackground(INTEL_HEADER_BG_KG).setFontColor(INTEL_HEADER_FG_KG).setFontWeight("bold");
  edges.setFrozenRows(1);
  [160, 190, 160, 90].forEach((w, i) => edges.setColumnWidth(i + 1, w));

  ui.alert("✅ Knowledge Graph Ready",
    (nodesNew ? "Created Knowledge_Graph. " : "Knowledge_Graph refreshed. ") +
    (edgesNew ? "Created Graph_Edges.\n\n" : "Graph_Edges refreshed.\n\n") +
    "Next: run syncCompanyNodes(), then generateGraphFromCompany() per company.",
    ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// SYNC — one Company node per Company_Master row (idempotent, keyed by Ref_ID)
// Company node IDs mirror the company ID: CO-0001 → N-CO-0001
// ══════════════════════════════════════════════════════════════════════════════
function syncCompanyNodes() {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const nodes  = ss.getSheetByName(SHEET_KG.NODES);
  if (!master || !nodes) { ui.alert("Run setup first — Company_Master or Knowledge_Graph missing."); return; }

  // Existing company nodes by Ref_ID
  const existing = new Set();
  nodes.getDataRange().getValues().slice(1).forEach(r => {
    if ((r[COL_KG_NODE.NODE_TYPE - 1] || "") === "Company") {
      existing.add((r[COL_KG_NODE.REF_ID - 1] || "").toString().trim());
    }
  });

  const mData = master.getDataRange().getValues();
  const newRows = [];
  for (let i = 1; i < mData.length; i++) {
    const coId = (mData[i][COL_COMPANY_MASTER.COMPANY_ID - 1] || "").toString().trim();
    const name = (mData[i][COL_COMPANY_MASTER.COMPANY_NAME - 1] || "").toString().trim();
    if (!coId || existing.has(coId)) continue;
    const seq = coId.replace(/^CO-/, "");
    newRows.push(["N-CO-" + seq, "Company", name, "", coId]);
  }

  if (newRows.length) nodes.getRange(nodes.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  ui.alert("✅ Company Nodes Synced", "Added " + newRows.length + " new company node(s).", ui.ButtonSet.OK);
}


// ══════════════════════════════════════════════════════════════════════════════
// GENERATE — Claude emits the people/regulator/industry/failure nodes + edges
// for one company, deduping nodes by (type, name).
// ══════════════════════════════════════════════════════════════════════════════
function generateGraphFromCompany(companyId) {
  const ss = intelSS_();
  const ui = SpreadsheetApp.getUi();

  if (!companyId) companyId = getSelectedCompanyId_();
  if (!companyId) return;

  const master = ss.getSheetByName(SHEET_COMPANY.MASTER);
  const mRow   = findCompanyRow_(master, companyId);
  if (mRow === -1) { ui.alert("Company_ID not found: " + companyId); return; }

  const nodes = ss.getSheetByName(SHEET_KG.NODES);
  const edges = ss.getSheetByName(SHEET_KG.EDGES);
  if (!nodes || !edges) { ui.alert("Run setupKnowledgeGraphTabs() first."); return; }

  const name     = master.getRange(mRow, COL_COMPANY_MASTER.COMPANY_NAME).getValue();
  const industry = master.getRange(mRow, COL_COMPANY_MASTER.INDUSTRY).getValue();
  const country  = master.getRange(mRow, COL_COMPANY_MASTER.COUNTRY).getValue();

  let profileCtx = "";
  const profile = ss.getSheetByName(SHEET_COMPANY.PROFILE);
  if (profile) {
    const pRow = findProfileRow_(profile, companyId);
    if (pRow !== -1) profileCtx = "Context: " + profile.getRange(pRow, COL_COMPANY_PROFILE.ROOT_CAUSE).getValue();
  }

  // Ensure the company itself is a node
  const companyNodeId = ensureNode_(nodes, "Company", name, "", companyId);

  const prompt = `
Extract the knowledge-graph entities and relationships for this company.

Company : ${name}
Industry: ${industry || "unknown"}
Country : ${country || "unknown"}
${profileCtx}

Emit NODES (entities connected to the company) and EDGES (relationships).
Node types allowed: Person, Board, Law, Regulator, Industry, Country, Framework.
(The company itself is already a node — reference it by its exact name "${name}".)

Keep names canonical and specific (e.g. "SEC", "Sarbanes-Oxley Act", not "the regulator").
Edge relationships e.g.: led_by, governed_by, regulated_by, operates_in, based_in,
failed_due_to, violated, uses_framework. Weight 0-100 = strength/centrality.

NODE_START
TYPE: [Person|Board|Law|Regulator|Industry|Country|Framework]
NAME: [canonical name]
DESC: [short description]
NODE_END

EDGE_START
SOURCE: [exact node name — usually "${name}"]
REL: [relationship]
TARGET: [exact node name]
WEIGHT: [0-100]
EDGE_END
`;

  try {
    const raw = callClaudeWithCustomSystem(prompt, INTEL_SYSTEM_CONTEXT, "high", 3000);

    // ── Nodes: ensure each, build name→node_id map (seed with the company) ──
    const nameToId = {};
    nameToId[normalizeName_(name)] = companyNodeId;

    const nodeBlocks = raw.match(/NODE_START([\s\S]*?)NODE_END/g) || [];
    let nodesAdded = 0;
    nodeBlocks.forEach(b => {
      const g = (f) => { const m = b.match(new RegExp(f + ":\\s*(.+)")); return m ? m[1].trim() : ""; };
      const type = g("TYPE"), nm = g("NAME"), desc = g("DESC");
      if (!KG_TYPE_CODE[type] || !nm) return;
      const before = nodes.getLastRow();
      const id = ensureNode_(nodes, type, nm, desc, "");
      nameToId[normalizeName_(nm)] = id;
      if (nodes.getLastRow() > before) nodesAdded++;
    });

    // ── Edges: resolve names to node IDs; skip edges with unknown endpoints ──
    const edgeBlocks = raw.match(/EDGE_START([\s\S]*?)EDGE_END/g) || [];
    const edgeRows = [];
    edgeBlocks.forEach(b => {
      const g = (f) => { const m = b.match(new RegExp(f + ":\\s*(.+)")); return m ? m[1].trim() : ""; };
      const src = nameToId[normalizeName_(g("SOURCE"))];
      const tgt = nameToId[normalizeName_(g("TARGET"))];
      const rel = g("REL");
      let w = parseInt(g("WEIGHT"), 10); if (isNaN(w)) w = 50;
      w = Math.max(0, Math.min(100, w));
      if (src && tgt && rel && !edgeExists_(edges, src, rel, tgt)) edgeRows.push([src, rel, tgt, w]);
    });
    if (edgeRows.length) edges.getRange(edges.getLastRow() + 1, 1, edgeRows.length, 4).setValues(edgeRows);

    master.getRange(mRow, COL_COMPANY_MASTER.UPDATED_AT).setValue(new Date());
    ui.alert("✅ Graph generated for " + name,
      nodesAdded + " new node(s), " + edgeRows.length + " new edge(s).", ui.ButtonSet.OK);

  } catch (err) {
    if (typeof logError === "function") logError("Intel 1.2 — Graph", companyId, "API/Runtime", err.message);
    ui.alert("❌ Graph generation failed: " + err.message);
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ══════════════════════════════════════════════════════════════════════════════
// Return existing node_id for (type,name) or create a new node and return its id.
function ensureNode_(nodes, type, name, desc, refId) {
  const data   = nodes.getDataRange().getValues();
  const target = normalizeName_(name);
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_KG_NODE.NODE_TYPE - 1] || "") === type &&
        normalizeName_(data[i][COL_KG_NODE.NODE_NAME - 1] || "") === target) {
      return (data[i][COL_KG_NODE.NODE_ID - 1] || "").toString().trim();
    }
  }
  const code = KG_TYPE_CODE[type] || "XX";
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const m = (data[i][COL_KG_NODE.NODE_ID - 1] || "").toString().match(new RegExp("^N-" + code + "-(\\d+)$"));
    if (m) { const s = parseInt(m[1], 10); if (s > maxSeq) maxSeq = s; }
  }
  const id = "N-" + code + "-" + String(maxSeq + 1).padStart(4, "0");
  nodes.appendRow([id, type, name, desc || "", refId || ""]);
  return id;
}

function edgeExists_(edges, src, rel, tgt) {
  const data = edges.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][COL_KG_EDGE.SOURCE_NODE - 1] || "").toString().trim() === src &&
        (data[i][COL_KG_EDGE.RELATIONSHIP - 1] || "").toString().trim() === rel &&
        (data[i][COL_KG_EDGE.TARGET_NODE - 1] || "").toString().trim() === tgt) return true;
  }
  return false;
}
