/* global console */

/* Operations parser — extract and validate structured operations from AI responses */

/**
 * Attempt to repair common LLM JSON mistakes.
 * Handles: trailing commas, missing closing brackets, single-line comments,
 * unescaped control chars in strings, trailing text after JSON.
 *
 * @param {string} raw - The raw JSON string that failed to parse
 * @returns {string|null} Repaired JSON string, or null if unrepairable
 */
function repairJSON(raw) {
  let s = raw.trim();

  // Strip single-line comments (// ...) outside of strings
  s = s.replace(/\/\/[^\n]*/g, "");

  // Strip block comments (/* ... */) outside of strings
  s = s.replace(/\/\*[\s\S]*?\*\//g, "");

  // Remove trailing commas before ] or }  —  the #1 LLM mistake
  // Handles: [1, 2, 3,] and {"a": 1, "b": 2,}
  s = s.replace(/,\s*([\]}])/g, "$1");

  // If the string ends mid-value (truncated), try to close open brackets
  const openBrackets = (s.match(/\[/g) || []).length;
  const closeBrackets = (s.match(/\]/g) || []).length;
  const openBraces = (s.match(/\{/g) || []).length;
  const closeBraces = (s.match(/\}/g) || []).length;

  // Close any unclosed string first (odd number of unescaped quotes)
  const unescapedQuotes = (s.match(/(?<!\\)"/g) || []).length;
  if (unescapedQuotes % 2 !== 0) {
    s += '"';
  }

  // Close missing brackets/braces
  for (let i = 0; i < openBrackets - closeBrackets; i++) s += "]";
  for (let i = 0; i < openBraces - closeBraces; i++) s += "}";

  // Strip trailing garbage after the last } or ] (models sometimes add text after JSON)
  const lastClose = Math.max(s.lastIndexOf("}"), s.lastIndexOf("]"));
  if (lastClose > 0 && lastClose < s.length - 1) {
    // Only strip if what follows looks like non-JSON text (not another valid char)
    const after = s.slice(lastClose + 1).trim();
    if (after.length > 0 && !/^[\s,:\d]/.test(after)) {
      s = s.slice(0, lastClose + 1);
    }
  }

  return s;
}

/**
 * Try to parse JSON with automatic repair fallback.
 * @param {string} text - Raw JSON string
 * @returns {{ parsed: object|null, repaired: boolean, error: string|null }}
 */
function safeParseJSON(text) {
  const trimmed = text.trim();

  // Fast path: try direct parse first
  try {
    const parsed = JSON.parse(trimmed);
    return { parsed, repaired: false, error: null };
  } catch (e) {
    // Attempt repair
    const repaired = repairJSON(trimmed);
    if (repaired && repaired !== trimmed) {
      try {
        const parsed = JSON.parse(repaired);
        console.log("[extract] JSON repaired successfully (original error:", e.message, ")");
        return { parsed, repaired: true, error: null };
      } catch (e2) {
        return { parsed: null, repaired: true, error: e2.message };
      }
    }
    return { parsed: null, repaired: false, error: e.message };
  }
}

/**
 * Extract structured operations JSON from AI response text.
 * Looks for a JSON code block (\`\`\`json ... \`\`\` or \`\`\` ... \`\`\`).
 * @param {string} response - Full AI response
 * @returns {{ plan: string, operations: object[], text: string, _switchMode: string|null, _switchMessage: string|null } | null}
 */
export function extractOperations(response) {
  if (!response) return null;

  // Try JSON code block first — [\r\n]* allows ```json{ without newline
  const jsonBlockMatch = response.match(/```(?:json)?\s*[\r\n]*([\s\S]*?)```/);
  if (jsonBlockMatch) {
    console.log("[extract] Found JSON code block, attempting parse...");
    const result = safeParseJSON(jsonBlockMatch[1]);
    if (result.parsed && typeof result.parsed === "object") {
      console.log(
        "[extract] Parsed JSON — plan:",
        result.parsed.plan ? result.parsed.plan.substring(0, 80) : "(none)",
        "ops:",
        result.parsed.operations?.length || 0,
        result.repaired ? "(repaired)" : ""
      );
      return {
        plan: result.parsed.plan || "",
        operations: Array.isArray(result.parsed.operations) ? result.parsed.operations : [],
        text: response,
        _switchMode: result.parsed._switchMode || null,
        _switchMessage: result.parsed._switchMessage || null,
      };
    }
    console.warn("[extract] JSON parse failed (even after repair):", result.error);
    // Log raw content around error position for debugging
    const rawContent = jsonBlockMatch[1].trim();
    const errorPos = result.error ? parseInt(result.error.match(/position (\d+)/)?.[1]) : -1;
    if (errorPos > 0) {
      const start = Math.max(0, errorPos - 80);
      const end = Math.min(rawContent.length, errorPos + 80);
      console.warn(
        "[extract] Raw JSON around error pos",
        errorPos,
        ":",
        JSON.stringify(rawContent.slice(start, end))
      );
    }
  }

  // Try generic code block — [\r\n]* allows ```{ without newline
  const blockMatch = response.match(/```\s*[\r\n]*([\s\S]*?)```/);
  if (blockMatch) {
    const result = safeParseJSON(blockMatch[1]);
    if (result.parsed && typeof result.parsed === "object") {
      console.log(
        "[extract] Parsed generic code block as JSON",
        result.repaired ? "(repaired)" : ""
      );
      return {
        plan: result.parsed.plan || "",
        operations: Array.isArray(result.parsed.operations) ? result.parsed.operations : [],
        text: response,
        _switchMode: result.parsed._switchMode || null,
        _switchMessage: result.parsed._switchMessage || null,
      };
    }
  }

  // No code block found — conversational response
  console.log("[extract] No code block found, treating as plain text");
  return { plan: "", operations: [], text: response.trim() };
}

/**
 * Validate operations against the provided registry.
 * @param {object[]} operations - Parsed operations
 * @param {object} registry - Operation registry map
 * @param {"interactive"|"explain"} mode
 * @returns {{ valid: object[], errors: string[] }}
 */
export function validateOperations(operations, registry, mode) {
  const valid = [];
  const errors = [];

  const explainOnlyOps = ["readRange"];

  // Field name aliases — models often use wrong names
  const FIELD_ALIASES = {
    sheet: "name",
    worksheet: "name",
    sheetName: "name",
    data: "values",
    formula: "formulas",
    fontName: "font",
    fontSize: "size",
    fill: "color",
    horizontal: "alignment",
    merged: "merge",
    bool: "bold",
  };

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const type = op.operation;
    const prevErrorCount = errors.length;

    // Normalize field aliases (only when canonical field is in the registry's required/optional fields)
    const entry2 = registry[type];
    if (entry2) {
      const knownFields = new Set(entry2.required || []);
      for (const [alias, canonical] of Object.entries(FIELD_ALIASES)) {
        if (alias in op && !(canonical in op) && knownFields.has(canonical)) {
          op[canonical] = op[alias];
          delete op[alias];
        }
      }
    }

    if (!type) {
      errors.push(`Operation ${i}: missing "operation" type`);
      continue;
    }

    // Block write operations in explain mode
    if (mode === "explain" && !explainOnlyOps.includes(type)) {
      errors.push(
        `Operation ${i} (${type}): not allowed in explain mode — switch to interactive mode for modifications`
      );
      continue;
    }

    const entry = registry[type];
    if (!entry) {
      errors.push(`Operation ${i}: unknown operation type "${type}"`);
      continue;
    }

    // Check required fields
    for (const field of entry.required) {
      if (op[field] === undefined || op[field] === null) {
        errors.push(`Operation ${i} (${type}): missing required field "${field}"`);
      }
    }

    // Run custom validator
    if (entry.validate && !entry.validate(op)) {
      errors.push(`Operation ${i} (${type}): failed validation — check field types`);
    }

    // Only push to valid if NO errors were added for this operation
    // (prevErrorCount avoids fragile "Operation ${i}:" string matching
    //  which would false-positive on "Operation 11" when checking index 1)
    if (errors.length === prevErrorCount) {
      valid.push(op);
    }
  }

  return { valid, errors };
}

const COMPACTION_THRESHOLD = 8;
const COMPACTION_KEEP = 4;
const COMPACTION_PROMPT = `You are an Excel task assistant. You have been executing a multi-step task.

Summarize your progress so far. Respond with a JSON object:
{
  "summary": "One-paragraph overview of what was done and what remains",
  "goal": "The original user request",
  "completed": ["list of what has been accomplished"],
  "errors": ["list of errors encountered and how they were fixed, or empty array"],
  "remaining": ["list of what still needs to be done, or empty array if complete"]
}

Respond with ONLY the JSON object, no explanation.`;

/**
 * Summarize conversation history to prevent context drift.
 * When history exceeds COMPACTION_THRESHOLD, calls the AI to produce a
 * structured self-summary (Cursor's "compaction-in-the-loop" pattern).
 * Falls back to deterministic truncation for shorter histories.
 * @param {Array} history - conversationHistory array
 * @param {function} [streamFromAI] - Optional async function to call the AI for self-summarization
 * @returns {Promise<Array>} Compressed message array
 */
export async function summarizeConversationHistory(history, streamFromAI) {
  if (!history || history.length <= 6) return history;

  // Keep the last N messages in full, summarize the rest
  const keep = history.slice(-COMPACTION_KEEP);
  const older = history.slice(0, -COMPACTION_KEEP);

  if (history.length < COMPACTION_THRESHOLD) {
    // Short history: use deterministic truncation (faster, no AI cost)
    return _deterministicSummary(older, keep);
  }

  // Long history: try AI self-summarization (Cursor "compaction-in-the-loop" pattern)
  if (streamFromAI) {
    try {
      const compacted = await _trySelfSummarize(older, keep, streamFromAI);
      if (compacted) return compacted;
    } catch (e) {
      console.warn(
        "[summarize] Self-summarization failed, falling back to deterministic:",
        e.message
      );
    }
  }

  // Fallback: deterministic truncation
  return _deterministicSummary(older, keep);
}

async function _trySelfSummarize(older, keep, streamFromAI) {
  const messagesForSummary = [
    { role: "system", content: COMPACTION_PROMPT },
    ...older.map((m) => ({
      role: m.role,
      content: typeof m.content === "string" ? m.content : JSON.stringify(m.content),
    })),
  ];

  const response = await streamFromAI(messagesForSummary, () => false, 1024);
  if (!response) return null;

  const jsonMatch = response.match(/```(?:json)?\s*[\r\n]*([\s\S]*?)```/);
  const rawText = jsonMatch ? jsonMatch[1].trim() : response.trim();
  const result = safeParseJSON(rawText);
  let parsed;
  if (result.parsed) {
    parsed = result.parsed;
  } else {
    console.warn("[summarize] Self-summary JSON parse failed, using raw text");
    parsed = { summary: response.substring(0, 500) };
  }

  const summaryText = `--- Conversation Compaction ---
${parsed.summary || "Conversation compressed."}
Goal: ${parsed.goal || "N/A"}
Completed: ${(parsed.completed || []).join("; ") || "None yet"}
Errors: ${(parsed.errors || []).join("; ") || "None"}
Remaining: ${(parsed.remaining || []).join("; ") || "N/A"}
--- End Compaction ---`;

  return [{ role: "system", content: summaryText }, ...keep];
}

function _deterministicSummary(older, keep) {
  const summaryLines = [];
  for (const msg of older) {
    if (msg.role === "user") {
      const short = msg.content.length > 200 ? msg.content.substring(0, 200) + "..." : msg.content;
      summaryLines.push(`  User: ${short}`);
    } else if (msg.role === "assistant") {
      const hasJson = /```json/.test(msg.content);
      if (hasJson) {
        const planMatch = msg.content.match(/"plan"\s*:\s*"([^"]+)"/);
        if (planMatch) {
          summaryLines.push(`  Assistant: [Plan] ${planMatch[1]}`);
        } else {
          summaryLines.push(`  Assistant: [Response] (truncated)`);
        }
      } else {
        const short =
          msg.content.length > 200 ? msg.content.substring(0, 200) + "..." : msg.content;
        summaryLines.push(`  Assistant: ${short}`);
      }
    }
  }

  const summary =
    summaryLines.length > 0
      ? `[Earlier conversation (${older.length} messages):\n${summaryLines.join("\n")}\n]`
      : "";

  return [{ role: "system", content: summary }, ...keep];
}

/**
 * Check if an AI response indicates task completion.
 */
export function isCompletionClaim(response) {
  return /complete\.$|done\.$|finished\.$|no more steps|all steps complete/i.test(response);
}

/**
 * Check if an AI/plain-text response is actionable (mentions operations).
 */
export function isActionableResponse(response) {
  return /operation|step|create|write|format|add|set|delete|chart|table/i.test(response);
}
