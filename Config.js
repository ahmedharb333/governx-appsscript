/* ============================================================================
   config.gs — GovernX Content OS
   Central configuration: constants, system prompt, legal & quality rules
   ============================================================================ */

// ── API Constants ─────────────────────────────────────────────────────────────
const ANTHROPIC_MODEL   = "claude-opus-4-6";
const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";

// ── Sheet Names ───────────────────────────────────────────────────────────────
const SHEET = {
  IDEA        : "Idea Catalogue",
  MASTER      : "Master Content Table",
  RESEARCH    : "Research Database",
  SCRIPT      : "Script Bank",
  VISUAL      : "Visual Library",
  PUBLISHING  : "Publishing Tracker",
  YOUTUBE     : "YouTube Metadata",
  DATA_MASTER : "Data_Master",
  ERROR_LOG   : "Error Log"
};

// ── Column Indices (1-based) ──────────────────────────────────────────────────

const COL_IDEA = {
  ID              : 1,
  COMPANY         : 2,
  DOMAIN          : 3,
  INDUSTRY        : 4,
  FIELD           : 5,
  TYPE            : 6,
  INITIAL_ANGLE   : 7,
  PRIORITY        : 8,
  LANGUAGE_FLAG   : 9,
  TARGET_FORMAT   : 10,
  CALL_TO_ACTION  : 11,
  SERIES          : 12,
  NOTE            : 13,
  PIPELINE_STATUS : 14   // ← N: auto-updated stage tracker e.g. "S1✅ S2✅ S3✅ S4⏳"
};

const COL_MASTER = {
  ID                     : 1,   // A
  TITLE                  : 2,   // B
  DOMAIN                 : 3,   // C
  INDUSTRY               : 4,   // D
  FIELD                  : 5,   // E
  TYPE                   : 6,   // F
  HOOK                   : 7,   // G
  FINAL_MOMENT           : 8,   // H
  REVERSE_ANGLE          : 9,   // I
  PRIMARY_ANGLE          : 10,  // J
  DISCIPLINE             : 11,  // K
  CORE_INSIGHT           : 12,  // L
  CHECKPOINTS            : 13,  // M
  TARGET_AUDIENCE        : 14,  // N
  LANGUAGE_FLAG          : 15,  // O
  SERIES                 : 16,  // P
  THUMBNAIL_BRIEF        : 17,  // Q ← new: thumbnail visual description
  NOTE                   : 18   // R ← shifted
};

const COL_RESEARCH = {
  ID             : 1,
  TOPIC          : 2,
  SOURCE_TYPE    : 3,
  DETAILS        : 4,
  SOURCE_LINK    : 5,
  KEY_INSIGHT    : 6,
  EVIDENCE_TYPE  : 7,
  TIMESTAMP      : 8,
  RELEVANCE      : 9,
  USED_IN_SCRIPT : 10,
  NOTE           : 11
};

const COL_SCRIPT = {
  ID                  : 1,
  ANALYSIS_FRAMEWORK  : 2,
  HOOK                : 3,
  NARRATIVE_FLOW      : 4,
  VOICEOVER_SCRIPT    : 5,
  GRC_BPR_CLOSING     : 6,
  SECTIONS            : 7,
  TARGET_FORMAT       : 8,
  DURATION            : 9,
  CALL_TO_ACTION      : 10,
  VERSION             : 11,
  STATUS              : 12,
  SERIES              : 13,
  NOTE                : 14
};

const COL_VISUAL = {
  ID               : 1,
  SCENE_NUM        : 2,
  SCENE_TYPE       : 3,
  DESCRIPTION      : 4,
  SOURCE           : 5,
  LINK             : 6,   // ← Search query / stock link — NEVER overwritten by script
  TIMESTAMP        : 7,
  LICENSE          : 8,
  STATUS           : 9,
  TASK_ID          : 10,  // ← KlingAI task ID
  AI_CLIP_URL      : 11,  // ← permanent Drive link to downloaded AI clip
  BUILT_WHERE      : 12,  // ← where to build this scene
  ASSEMBLY_NOTES   : 13,  // ← exact CapCut instruction for this scene
  VEO_PROMPT       : 14,  // ← formatted Google Veo / AI Studio prompt
  CHECKPOINT_DATE  : 15,  // ← year/date for Checkpoint scenes
  CHECKPOINT_EVENT : 16,  // ← what happened at this checkpoint
  CHECKPOINT_ANGLE : 17,  // ← governance/risk/process angle in caps
  NOTE             : 18
};

// ── Visual Library Extended Columns (Stage 4B Director Review output) ────────
const COL_VISUAL_EXTENDED = {
  REMOTION_DATA        : 19,  // Structured key=value string for Remotion renderer
  REMOTION_STYLE       : 20,  // Template name, variant, weight
  VOICEOVER_SYNC       : 21,  // Exact voiceover sentence(s) this scene covers
  SCENE_SCORE          : 22,  // A | B | NEEDS_DATA | TEMPLATE_REQUEST
  VOICEOVER_AUDIO_URL  : 23   // Drive URL for per-scene ElevenLabs MP3 (Stage 7B)
};

const COL_PUBLISHING = {
  ID                 : 1,
  PUBLISH_DATE       : 2,
  TITLE_FINAL        : 3,
  THUMBNAIL          : 4,
  VIEWS              : 5,
  CTR                : 6,
  RETENTION          : 7,
  AVG_WATCH_TIME     : 8,
  PERF_FORMULA       : 9,
  PERF_SCORE         : 10,
  SCORE              : 11,
  REPEAT_DECISION    : 12,
  PRODUCTION_PACKAGE : 13,  // ← 📄 Open Package
  VOICEOVER_AUDIO    : 14,  // ← 🎵 Play Audio
  SCENES_FOLDER      : 15,  // ← 📁 Open Scenes
  ASSEMBLY_GUIDE     : 16,  // ← 📋 Open Guide
  NOTES              : 17,  // ← benchmarks and manual notes
  YOUTUBE_VIDEO_ID   : 18   // ← paste YouTube video ID here after upload (e.g. dQw4w9WgXcQ)
};

// ── YouTube Metadata Tab Columns ──────────────────────────────────────────────
const COL_YOUTUBE = {
  ID             : 1,
  TITLE_A        : 2,   // Primary title
  TITLE_B        : 3,   // A/B test variant
  TITLE_C        : 4,   // A/B/C test variant (question format)
  DESCRIPTION    : 5,   // Full YouTube description
  TAGS           : 6,   // Comma-separated tags
  CHAPTERS       : 7,   // Timestamp chapters
  HASHTAGS       : 8,   // #hashtags
  FIRST_COMMENT  : 9,   // Pinned first comment
  END_SCREEN     : 10,  // End screen CTA text
  THUMBNAIL_BRIEF: 11,  // Visual brief for thumbnail design
  STATUS         : 12,  // Draft / Ready
  NOTE           : 13
};


const COL_ERROR = {
  TIMESTAMP  : 1,
  STAGE      : 2,
  ID         : 3,
  ERROR_TYPE : 4,
  DETAILS    : 5,
  RESOLVED   : 6
};




const DOMAIN_CODES = {
  "Business"                : "BIZ",
  "Sports"                  : "SPT",
  "Media & Creator Economy" : "MED",
  "Public Sector"           : "PUB",
  "Startups & Tech"         : "TECH"
};

// ── Google Drive folder structure ─────────────────────────────────────────────
const DRIVE_FOLDER_NAME    = "GovernX Production Packages";
// Subfolders created per content ID: DRIVE_FOLDER_NAME / GX-XXXX-XXX-001 / ...
const DRIVE_SCENES_SUBFOLDER = "Scenes";

// ── ElevenLabs API ────────────────────────────────────────────────────────────
const ELEVENLABS_API_URL   = "https://api.elevenlabs.io/v1/text-to-speech";
const ELEVENLABS_MODEL     = "eleven_multilingual_v2";

// ── Pexels API ────────────────────────────────────────────────────────────────
const PEXELS_API_URL       = "https://api.pexels.com/videos/search";

// ── Google Veo / Gemini API ───────────────────────────────────────────────────
const GEMINI_API_BASE      = "https://generativelanguage.googleapis.com/v1beta";
const VEO_MODEL            = "veo-2.0-generate-001"; // Veo 3 requires special access — use Veo 2
const VEO_POLL_INTERVAL    = 20000;  // 20 seconds between polls
const VEO_MAX_POLLS        = 30;     // max 10 min wait per clip
const KLING_API_URL        = "https://api.klingai.com/v1/videos/text2video";
const KLING_API_STATUS_URL = "https://api.klingai.com/v1/videos/text2video";
const KLING_MODEL          = "kling-v2";
const KLING_POLL_INTERVAL  = 15000;  // 15 sec between status checks
const KLING_MAX_POLLS      = 40;     // max 10 min wait per clip

// ── Performance Benchmark Thresholds ─────────────────────────────────────────
const BENCHMARKS = {
  CTR_SCALE      : 6,    // CTR % to trigger Scale decision
  RETENTION_SCALE: 45,   // Retention % to trigger Scale
  VIEWS_KILL     : 500   // Views at 30 days to trigger Kill
};

// ── Word Count Ranges by Format ───────────────────────────────────────────────
const WORD_COUNT = {
  "Short (< 90s)"      : { min: 150,  max: 250  },
  "Standard (4–7 min)" : { min: 600,  max: 900  },
  "Deep Dive (10–15 min)": { min: 1500, max: 2000 }
};

// ── Legal Guardrails & Quality Rules (defined FIRST — used inside SYSTEM_CONTEXT) ──
const LEGAL_AND_QUALITY_RULES = `
=== LEGAL GUARDRAILS — NON-NEGOTIABLE ===

1. NO UNVERIFIED CLAIMS: Only state what is supported by the sources provided. 
   No speculation presented as fact. If uncertain, use "evidence suggests" or 
   "reported decisions indicate."

2. ATTRIBUTION REQUIRED: Every factual claim in the script must be traceable 
   to a source in the Research Database provided to you.

3. NO DEFAMATORY LANGUAGE: Individuals are described by their decisions and 
   organizational roles only. No personal attacks, character assassinations, 
   or accusations beyond what sources support.

4. LIVING PERSONS FLAG: If the story involves living executives or public figures, 
   use cautious framing — "reported," "according to," "evidence suggests." 
   Never state opinions as facts about living individuals.

5. FAIR USE FRAMING: All referenced content (footage suggestions, quotes, memos) 
   must be framed as commentary and critical analysis. This protects the channel 
   under fair use doctrine.

6. NO FABRICATED QUOTES: Never invent or paraphrase dialogue as if it were a 
   direct quote. Direct quotes must come verbatim from verified sources only. 
   Use: Source says: "exact words" — not reconstructed versions.

7. COPYRIGHT FLAG: If any suggested scene or source carries copyright risk, 
   flag it explicitly with [COPYRIGHT RISK] in your output.

=== QUALITY STANDARDS — GOVERNX FORMAT ===

STRUCTURE (mandatory in every script):
- Open with the visible outcome — a specific date, number, or event. Never vague.
- Follow the reverse engineering path: Outcome → Decision Chain → Root Cause → Lesson
- Close with an explicit GRC or BPR closing argument — named, not implied

SCRIPT QUALITY:
- No filler language: remove "it is worth noting," "interestingly," "as we can see," 
  "it goes without saying," "needless to say"
- Every checkpoint must be a specific moment — not a general observation
- The hook must create urgency in the first 5 seconds

EVIDENCE STANDARD:
- Reference a minimum of 4 sources from the Research Database in the script
- At the end of your script output, include a QUALITY CHECK block in this exact format:

QUALITY_CHECK_START
sources_referenced: [number]
reverse_engineering_structure: YES/NO
grc_bpr_closing_argument: YES/NO
hook_opens_with_specific_fact: YES/NO
word_count: [number]
living_persons_flagged: YES/NO
sources_used: [comma-separated list of source Details fields used]
copyright_risks_flagged: YES/NO
QUALITY_CHECK_END
`;

// ── GovernX System Context (defined AFTER LEGAL_AND_QUALITY_RULES) ────────────
const SYSTEM_CONTEXT =
`You are the content engine for GovernX — a YouTube channel that reverse-engineers 
leadership decisions to reveal the hidden role of GRC (Governance, Risk & Compliance) 
and BPR (Business Process Reengineering) in every major organizational success or failure.

YOUR CORE METHOD — REVERSE ENGINEERING:
1. Start with the visible outcome (the headline the world saw)
2. Identify the decision chain (what leaders did or failed to do)
3. Trace back to the root cause (governance gap, risk blind spot, or broken process)
4. Land the GRC/BPR lesson (what proper governance or process reengineering would have changed)

YOUR TONE:
- Sharp, insightful, built for video storytelling
- Narrative-driven — not academic, not bullet-point dry
- The audience feels the stakes before they understand the lesson
- GRC and BPR are never presented as compliance checkboxes — they are the hidden engine 
  behind every great or catastrophic organizational outcome

YOUR AUDIENCE:
Business leaders, C-Suite executives, board members, managers, compliance professionals, 
and business students who need to understand WHY organizations succeed or fail at a 
systemic level — not just the surface story.

OUTPUT LANGUAGE RULES — APPLY EXACTLY AS SPECIFIED:

If Language Flag = English:
- Write ALL fields in English

If Language Flag = Arabic:
- Write these fields in Modern Standard Arabic (فصحى): 
  TITLE, HOOK, FINAL_MOMENT, CORE_INSIGHT, VOICEOVER_SCRIPT, GRC_BPR_CLOSING, 
  THUMBNAIL_BRIEF, CHECKPOINT_DATE, CHECKPOINT_EVENT, CHECKPOINT_ANGLE,
  YouTube TITLE_A, TITLE_B, DESCRIPTION, HASHTAGS, FIRST_COMMENT
- Write these fields in English (technical/structural):
  DOMAIN, INDUSTRY, FIELD, TYPE, DISCIPLINE, REVERSE_ANGLE, PRIMARY_ANGLE,
  TARGET_AUDIENCE, LANGUAGE_FLAG, SERIES, SECTIONS, VERSION, STATUS,
  Scene SOURCE, LICENSE, BUILT_WHERE, VEO_PROMPT, TAGS
- Arabic text must be right-to-left, professional register, no dialect
- Numbers and statistics stay as numerals (2013, $7.2B) — do not write in Arabic words

If Language Flag = Bilingual:
- VOICEOVER_SCRIPT and GRC_BPR_CLOSING → Arabic
- All other fields → English
- Hook → write in BOTH languages, Arabic first then English below

` + LEGAL_AND_QUALITY_RULES;