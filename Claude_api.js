/* ============================================================================
   claude_api.gs — GovernX Content OS
   Claude API caller with retry logic, rate limit handling, error logging

   VERSION 2.0 UPGRADES:
   + Adaptive thinking with per-stage effort levels (low / medium / high / max)
   + Prompt caching on SYSTEM_CONTEXT (saves 60–70% on input token costs)
   + Cache hit logging for cost monitoring
   + callClaude() now accepts an optional effort parameter
   ============================================================================ */


// ── Per-stage effort map ──────────────────────────────────────────────────────
// Controls how deeply Claude reasons per stage.
// "high" = default — thinks on most queries.
// "max"  = deepest reasoning — use for root cause analysis + script generation.
// "medium" = balanced — use for structured output stages.
// "low"  = fastest + cheapest — use for metadata and formatting stages.
//
// Reference: https://platform.claude.com/docs/en/build-with-claude/adaptive-thinking
//
const STAGE_EFFORT = {
  "stage_1_master"    : "high",    // Root cause analysis — high quality, respects 6-min Apps Script limit
  "stage_2_research"  : "high",    // Source identification — benefit from thinking
  "stage_3_script"    : "medium",  // Script generation — "high" + adaptive thinking + 32k max_tokens ran non-streaming past the UrlFetchApp/Anthropic HTTP timeout on Opus 4.6, then retried 3× and blew the 6-min Apps Script cap ("Exceeded maximum execution time"). "medium" returns in time; Stage 3B QA Review refines quality after.
  "stage_4_scenes"    : "low",     // Scene routing — structured (cut script into scene blocks). At "medium", adaptive thinking on a 26-30 scene Standard video ran past the 6-min Apps Script limit → timeout. Low thinking keeps the single call fast enough; 4B refines quality after.
  "stage_4_director"  : "medium",  // Director review — structured data fill; high-effort thinking eats the output budget and truncates many-scene videos (was falling through to "high")
  "stage_5_publishing": "low",     // Row creation — formatting task
  "stage_6_package"   : "low",     // Document export — formatting task
  "stage_7_voiceover" : "low",     // ElevenLabs prep — formatting task
  "stage_8_scenes"    : "medium",  // Scene processing — some judgment needed
  "stage_9_assembly"  : "medium",  // Assembly guide — structured output
  "stage_10_youtube"  : "medium",  // Metadata — SEO judgment benefits from thinking
  "default"           : "high"     // Fallback for any unspecified stage
};


// ══════════════════════════════════════════════════════════════════════════════
// callClaude() — PRIMARY API CALLER
//
// USAGE:
//   callClaude(prompt)                    → uses "high" effort (default)
//   callClaude(prompt, "stage_3_script")  → uses "max" effort from STAGE_EFFORT map
//   callClaude(prompt, "stage_1_master")  → uses "max" effort
//   callClaude(prompt, "stage_5_publishing") → uses "low" effort
//
// WHAT CHANGED FROM v1:
//   1. system is now an ARRAY with cache_control — enables prompt caching
//   2. thinking: {type: "adaptive"} + output_config: {effort: X} — adaptive thinking
//   3. max_tokens raised to 16000 — Opus 4.6 supports 128k output; 16k covers Deep Dive scripts
//   4. Cache hit stats logged to Logger for cost monitoring
// ══════════════════════════════════════════════════════════════════════════════

function callClaude(finalPrompt, stageKey) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("ANTHROPIC_API_KEY");

  if (!apiKey) throw new Error("ANTHROPIC_API_KEY missing from Script Properties");

  // ── Resolve effort level for this stage ──────────────────────────────────
  const effort = STAGE_EFFORT[stageKey] || STAGE_EFFORT["default"];

  const MAX_TRIES     = 3;
  const RETRY_WAIT_MS = 6000;
  const RATELIMIT_MS  = 15000;
  let   lastError     = "";

  for (let attempt = 1; attempt <= MAX_TRIES; attempt++) {

    try {

      // Stage 3 needs more tokens: voiceover + SCENE_BLUEPRINT + DATA_MOMENTS
      // Standard 16k is insufficient when the prompt itself consumes 6-8k tokens
      // — or when the output is large. The director review emits REMOTION_DATA for
      // EVERY scene; a 4-7 min video with ~20 scenes needs the bigger budget or the
      // response truncates and zero scenes parse. Stage 4 (scenes) has the SAME
      // problem: adaptive thinking + 26-30 scene blocks overrun 16k → 0 scenes.
      const stageMaxTokens = (stageKey === "stage_3_script" || stageKey === "stage_4_director" || stageKey === "stage_4_scenes") ? 32000 : 16000;

      const payload = {
        model      : ANTHROPIC_MODEL,
        max_tokens : stageMaxTokens,

        // ── Prompt caching — system as array with cache_control ─────────────
        // SYSTEM_CONTEXT + LEGAL_AND_QUALITY_RULES is ~4500+ tokens.
        // Marked with cache_control so it is cached after first call.
        // Cache TTL: 5 minutes (resets on every hit).
        // Cost: first call pays 1.25x write; subsequent calls pay 0.1x read.
        // For a 10-stage pipeline run: ~9 out of 10 calls read from cache.
        system: [
          {
            type         : "text",
            text         : SYSTEM_CONTEXT,
            cache_control: { type: "ephemeral" }
          }
        ],

        // ── Adaptive thinking ────────────────────────────────────────────────
        // Replaces deprecated budget_tokens approach.
        // Claude decides when and how much to think based on effort level.
        // "max"  → thinks deeply on almost every query
        // "high" → thinks on most queries (default)
        // "medium" → selective thinking
        // "low"  → minimal thinking — fast and cheap for formatting tasks
        thinking: {
          type: "adaptive"
        },

        output_config: {
          effort: effort
        },

        messages: [
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

        // ── Log cache stats for cost monitoring ──────────────────────────────
        if (json.usage) {
          const cacheWrite = json.usage.cache_creation_input_tokens || 0;
          const cacheRead  = json.usage.cache_read_input_tokens     || 0;
          const freshInput = json.usage.input_tokens                || 0;
          const stage      = stageKey || "default";

          Logger.log(
            "[GovernX API] Stage: " + stage +
            " | Effort: " + effort +
            " | Fresh: " + freshInput +
            " | Cache write: " + cacheWrite +
            " | Cache read: " + cacheRead +
            (cacheRead > 0 ? " ✅ CACHE HIT" : " ⚡ cache miss")
          );
        }

        // ── Extract text from content blocks ─────────────────────────────────
        // With adaptive thinking enabled, response may contain thinking blocks
        // and text blocks. We want only the text content.
        const textContent = json.content
          .filter(block => block.type === "text")
          .map(block => block.text)
          .join("");

        return textContent;
      }

      if (code === 529) {
        lastError = "API overloaded (529)";
        Logger.log("Attempt " + attempt + "/" + MAX_TRIES +
          ": overloaded. Waiting " + RETRY_WAIT_MS / 1000 + "s…");
        if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
        continue;
      }

      if (code === 429) {
        lastError = "Rate limited (429)";
        Logger.log("Attempt " + attempt + "/" + MAX_TRIES +
          ": rate limited. Waiting " + RATELIMIT_MS / 1000 + "s…");
        if (attempt < MAX_TRIES) Utilities.sleep(RATELIMIT_MS);
        continue;
      }

      // ── Claude credit exhausted → automatic Gemini fallback ───────────────
      // A zero-balance 400 will NEVER succeed on retry, so don't spend the other
      // attempts on it — go straight to the free provider and return its text.
      if (code === 400 && /credit balance is too low/i.test(body)) {
        const fb = geminiFallback_(SYSTEM_CONTEXT, finalPrompt, stageMaxTokens, stageKey);
        if (fb !== null) return fb;
      }

      throw new Error("Claude API error " + code + ": " + body);

    } catch (err) {
      lastError = err.message;
      Logger.log("Attempt " + attempt + " failed: " + err.message);
      // A UrlFetchApp "Timeout" means the generation was too slow to return, not a
      // transient blip. Retrying the identical call just times out again and burns
      // the 6-min Apps Script budget. So fall straight to the free Gemini provider
      // (fast Flash — it won't itself time out) if a key is set; otherwise fail with
      // an actionable message rather than retrying into another timeout.
      if (/timeout|timed out|deadline|took too long/i.test(err.message)) {
        const fb = geminiFallback_(SYSTEM_CONTEXT, finalPrompt, 16000, stageKey);
        if (fb !== null) return fb;
        throw new Error(
          "Claude API timed out for stage '" + (stageKey || "default") + "' (effort: " + effort +
          "). Too slow to return non-streaming, and no GEMINI_API_KEY is set for fallback. Add a " +
          "Gemini key (free), lower this stage's effort in STAGE_EFFORT, or reduce its max_tokens."
        );
      }
      if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
    }
  }

  throw new Error(
    "Claude API failed after " + MAX_TRIES +
    " attempts. Last error: " + lastError
  );
}


// ══════════════════════════════════════════════════════════════════════════════
// callClaudeWithCustomSystem() — FOR SKILLS THAT USE THEIR OWN SYSTEM CONTEXT
//
// Used by:
//   - Director_Skill.gs   → callClaudeAsDirector()
//   - Skills_Library.gs   → callClaudeAsEvaluator()
//   - Company_Selector_Skill.gs → callClaudeAsSelector()
//
// These functions build their own payloads with their own system contexts.
// This helper is provided as a convenience wrapper with the same retry logic
// and cache support, but accepting a custom system context string.
//
// USAGE:
//   callClaudeWithCustomSystem(prompt, MY_CUSTOM_SYSTEM, "max", 12000)
// ══════════════════════════════════════════════════════════════════════════════

function callClaudeWithCustomSystem(finalPrompt, systemContext, effort, maxTokens) {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("ANTHROPIC_API_KEY");

  if (!apiKey) throw new Error("ANTHROPIC_API_KEY missing from Script Properties");

  const resolvedEffort    = effort    || "high";
  const resolvedMaxTokens = maxTokens || 8000;

  const MAX_TRIES     = 3;
  const RETRY_WAIT_MS = 8000;
  const RATELIMIT_MS  = 20000;
  let   lastError     = "";

  for (let attempt = 1; attempt <= MAX_TRIES; attempt++) {
    try {
      const payload = {
        model      : ANTHROPIC_MODEL,
        max_tokens : resolvedMaxTokens,
        system: [
          {
            type         : "text",
            text         : systemContext,
            cache_control: { type: "ephemeral" }  // cache custom system contexts too
          }
        ],
        thinking: {
          type: "adaptive"
        },
        output_config: {
          effort: resolvedEffort
        },
        messages: [{ role: "user", content: finalPrompt }]
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

        if (json.usage) {
          Logger.log(
            "[GovernX Custom] Effort: " + resolvedEffort +
            " | Fresh: "      + (json.usage.input_tokens                || 0) +
            " | Cache write: " + (json.usage.cache_creation_input_tokens || 0) +
            " | Cache read: "  + (json.usage.cache_read_input_tokens     || 0)
          );
        }

        return json.content
          .filter(block => block.type === "text")
          .map(block => block.text)
          .join("");
      }

      if (code === 529 || code === 429) {
        lastError = "API error " + code;
        Utilities.sleep(code === 429 ? RATELIMIT_MS : RETRY_WAIT_MS);
        continue;
      }

      if (code === 400 && /credit balance is too low/i.test(body)) {
        const fb = geminiFallback_(systemContext, finalPrompt, resolvedMaxTokens, "custom");
        if (fb !== null) return fb;
      }

      throw new Error("API error " + code + ": " + body);

    } catch (err) {
      lastError = err.message;
      if (/timeout|timed out|deadline|took too long/i.test(err.message)) {
        const fb = geminiFallback_(systemContext, finalPrompt, resolvedMaxTokens, "custom");
        if (fb !== null) return fb;
      }
      if (attempt < MAX_TRIES) Utilities.sleep(RETRY_WAIT_MS);
    }
  }

  throw new Error("API failed after " + MAX_TRIES + " attempts. Last: " + lastError);
}


// ══════════════════════════════════════════════════════════════════════════════
// geminiFallback_() — FREE AUTOMATIC FALLBACK when Claude is unavailable
//
// Fires ONLY when Claude returns "credit balance is too low" (console at $0) or a
// UrlFetchApp timeout. Keeps the whole pipeline flowing on Google Gemini's genuinely
// free tier (a real per-account quota, not a shared pool) while the Anthropic
// console is unfunded. Shared by every GovernX Claude entry point:
//   callClaude, callClaudeWithCustomSystem, callClaudeAsDirector,
//   callClaudeAsEvaluator, callClaudeAsSelector.
//
// TRIGGER MODE: auto only. There is no manual switch — Claude is always tried first;
// Gemini runs only when Claude cannot. When funded, this code never executes.
//
// QUALITY NOTE: Gemini Flash is below Claude, especially on adversarial reasoning.
// Every Gemini run is logged to the GEMINI_FALLBACK_LOG Script Property so you can
// re-run those stages on Claude once the console is funded. Draft-while-broke,
// finalise-on-Claude-when-funded.
//
// Returns the model's text, or null when no GEMINI_API_KEY is set (so the caller
// raises its normal error instead of silently doing nothing).
//
// Uses Gemini's OpenAI-compatible endpoint, so the request shape matches every
// other OpenAI-style provider. Model overridable via the GEMINI_MODEL property
// (default gemini-2.0-flash — solidly free; set gemini-2.5-flash for higher quality
// if your key has it).
// ══════════════════════════════════════════════════════════════════════════════
function geminiFallback_(systemText, userPrompt, maxTokens, stageKey) {
  const props = PropertiesService.getScriptProperties();
  const key   = props.getProperty("GEMINI_API_KEY");
  if (!key) return null;   // no key configured → caller throws its normal error

  const model = props.getProperty("GEMINI_MODEL") || "gemini-2.0-flash";
  const url   = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions";

  const resp = UrlFetchApp.fetch(url, {
    method            : "post",
    contentType       : "application/json",
    headers           : { "Authorization": "Bearer " + key },
    payload           : JSON.stringify({
      model,
      // Gemini 2.0 Flash caps output at 8192 tokens; every GovernX stage's real
      // output (a script, a director batch, a metadata block) fits well inside it.
      max_tokens: Math.min(maxTokens || 8000, 8192),
      messages  : [
        { role: "system", content: String(systemText || "") },
        { role: "user",   content: String(userPrompt  || "") }
      ]
    }),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const body = resp.getContentText();
  if (code !== 200) {
    throw new Error("Gemini fallback (" + model + ") failed " + code + ": " + body.substring(0, 300));
  }

  const json = JSON.parse(body);
  const text = (json.choices && json.choices[0] && json.choices[0].message &&
                json.choices[0].message.content) || "";

  // Flag every fallback run so you know exactly which stages to re-run on Claude.
  const stamp = new Date().toISOString() + " — " + (stageKey || "custom") +
                " ran on " + model + " (re-run on Claude when funded)";
  const log = props.getProperty("GEMINI_FALLBACK_LOG") || "";
  props.setProperty("GEMINI_FALLBACK_LOG", (log + "\n" + stamp).slice(-4000));
  Logger.log("⚠ [GovernX] Claude unavailable → Gemini fallback '" + model +
             "' used for '" + (stageKey || "custom") + "'. Quality below Claude — " +
             "re-run this stage on Claude when the console is funded.");
  return text;
}


// ══════════════════════════════════════════════════════════════════════════════
// PIPELINE.GS INTEGRATION — STAGE KEY REFERENCE
//
// To activate per-stage effort, pass the stage key as the second argument
// wherever callClaude() is called in Pipeline.gs:
//
//   Stage 1:  callClaude(prompt, "stage_1_master")
//   Stage 2:  callClaude(prompt, "stage_2_research")
//   Stage 3:  callClaude(prompt, "stage_3_script")
//   Stage 4:  callClaude(prompt, "stage_4_scenes")
//   Stage 5:  callClaude(prompt, "stage_5_publishing")
//   Stage 6:  callClaude(prompt, "stage_6_package")
//   Stage 7:  callClaude(prompt, "stage_7_voiceover")
//   Stage 8:  callClaude(prompt, "stage_8_scenes")
//   Stage 9:  callClaude(prompt, "stage_9_assembly")
//   Stage 10: callClaude(prompt, "stage_10_youtube")
//
// Each is a one-word addition to an existing callClaude() call.
// No other changes to Pipeline.gs are needed for the API upgrade.
//
// DIRECTOR + EVALUATOR + SELECTOR:
// Those three skills already have their own callClaude* functions
// (callClaudeAsDirector, callClaudeAsEvaluator, callClaudeAsSelector).
// Optionally refactor them to use callClaudeWithCustomSystem() instead
// to eliminate duplicate retry logic — but not required.
// ══════════════════════════════════════════════════════════════════════════════