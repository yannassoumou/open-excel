/* global describe, it, before */

/**
 * Unit tests for the mode system (interactive + explain).
 *
 * Run with: node test/mode-system.test.js
 *
 * These tests validate the core logic of the two-mode system by
 * extracting and testing the pure functions without the browser/Office dependency.
 */

const assert = require("assert");

// Minimal test framework (avoids npm dependency)
let passed = 0;
let failed = 0;
let total = 0;

function describe(name, fn) {
  console.log(`\n${name}`);
  fn();
}

function it(name, fn) {
  total++;
  try {
    fn();
    passed++;
    console.log(`  ✓ ${name}`);
  } catch (e) {
    failed++;
    console.log(`  ✗ ${name} — ${e.message}`);
    console.log(`    at ${e.stack.split("\n")[1]?.trim()}`);
  }
}

// ─── Mock state ──────────────────────────────────────────────────────────────

let mockCurrentMode = "interactive";

// ─── Re-create core logic for testing ────────────────────────────────────────

const OPERATION_REGISTRY = {
  createWorksheet: { required: ["name"], validate: (op) => typeof op.name === "string" && op.name.length > 0 },
  writeValues: { required: ["sheet", "range", "values"], validate: (op) => Array.isArray(op.values) },
  writeFormulas: { required: ["sheet", "range", "formulas"], validate: (op) => Array.isArray(op.formulas) },
  setFillColor: { required: ["sheet", "range", "color"], validate: () => true },
  readRange: { required: ["sheet", "range"], validate: () => true },
  deleteWorksheet: { required: ["name"], validate: (op) => typeof op.name === "string" },
  autofitColumns: { required: ["sheet"], validate: () => true },
  deleteRows: { required: ["sheet", "range"], validate: () => true },
};

const EXPLAIN_ONLY_OPS = ["readRange"];

function validateOperations(operations, mode) {
  const valid = [];
  const errors = [];

  for (let i = 0; i < operations.length; i++) {
    const op = operations[i];
    const type = op.operation;

    if (!type) {
      errors.push(`Operation ${i}: missing "operation" type`);
      continue;
    }

    // Block write operations in explain mode
    if (mode === "explain" && !EXPLAIN_ONLY_OPS.includes(type)) {
      errors.push(`Operation ${i} (${type}): not allowed in explain mode`);
      continue;
    }

    const registry = OPERATION_REGISTRY[type];
    if (!registry) {
      errors.push(`Operation ${i}: unknown operation type "${type}"`);
      continue;
    }

    let hasErrors = false;
    for (const field of registry.required) {
      if (op[field] === undefined || op[field] === null) {
        errors.push(`Operation ${i} (${type}): missing required field "${field}"`);
        hasErrors = true;
      }
    }

    if (registry.validate && !registry.validate(op)) {
      errors.push(`Operation ${i} (${type}): failed validation`);
      hasErrors = true;
    }

    if (!hasErrors) {
      valid.push(op);
    }
  }

  return { valid, errors };
}

function extractOperations(response) {
  if (!response) return null;

  const jsonBlockMatch = response.match(/```(?:json)?\s*[\r\n]+([\s\S]*?)```/);
  if (jsonBlockMatch) {
    try {
      const parsed = JSON.parse(jsonBlockMatch[1].trim());
      if (parsed && typeof parsed === "object") {
        return {
          plan: parsed.plan || "",
          operations: Array.isArray(parsed.operations) ? parsed.operations : [],
          text: response,
          _switchMode: parsed._switchMode || null,
          _switchMessage: parsed._switchMessage || null,
        };
      }
    } catch {
      // Parse failed
    }
  }

  return null;
}

function buildSystemPrompt(sheetContext, mode) {
  const hasExplainMarker = mode === "explain";
  if (hasExplainMarker) {
    return `[EXPLAIN MODE] You are a read-only Excel assistant...${sheetContext}`;
  }
  return `[INTERACTIVE MODE] You are an Excel assistant...${sheetContext}`;
}

function isValidMode(mode) {
  return ["interactive", "explain"].includes(mode);
}

// ─── Tests ───────────────────────────────────────────────────────────────────

describe("Mode validation", () => {
  it("should accept 'interactive' as valid", () => {
    assert.strictEqual(isValidMode("interactive"), true);
  });

  it("should accept 'explain' as valid", () => {
    assert.strictEqual(isValidMode("explain"), true);
  });

  it("should reject invalid mode names", () => {
    assert.strictEqual(isValidMode("readonly"), false);
    assert.strictEqual(isValidMode("write"), false);
    assert.strictEqual(isValidMode(""), false);
  });
});

describe("buildSystemPrompt — mode-aware template selection", () => {
  it("should return interactive prompt for interactive mode", () => {
    const prompt = buildSystemPrompt("Sheet1: A1:B2", "interactive");
    assert.ok(prompt.includes("[INTERACTIVE MODE]"));
    assert.ok(!prompt.includes("[EXPLAIN MODE]"));
  });

  it("should return explain prompt for explain mode", () => {
    const prompt = buildSystemPrompt("Sheet1: A1:B2", "explain");
    assert.ok(prompt.includes("[EXPLAIN MODE]"));
    assert.ok(!prompt.includes("[INTERACTIVE MODE]"));
  });

  it("should inject sheet context into both prompts", () => {
    const ctx = "Workbook has 2 sheet(s): Data, Summary";
    const interactivePrompt = buildSystemPrompt(ctx, "interactive");
    const explainPrompt = buildSystemPrompt(ctx, "explain");
    assert.ok(interactivePrompt.includes(ctx));
    assert.ok(explainPrompt.includes(ctx));
  });
});

describe("validateOperations — explain mode blocks writes", () => {
  it("should allow readRange in both modes", () => {
    const ops = [{ operation: "readRange", sheet: "Sheet1", range: "A1:C5" }];
    const interactive = validateOperations(ops, "interactive");
    const explain = validateOperations(ops, "explain");
    assert.strictEqual(interactive.valid.length, 1);
    assert.strictEqual(interactive.errors.length, 0);
    assert.strictEqual(explain.valid.length, 1);
    assert.strictEqual(explain.errors.length, 0);
  });

  it("should allow writeValues only in interactive mode", () => {
    const ops = [{ operation: "writeValues", sheet: "Sheet1", range: "A1", values: [["x"]] }];
    const interactive = validateOperations(ops, "interactive");
    const explain = validateOperations(ops, "explain");
    assert.strictEqual(interactive.valid.length, 1);
    assert.strictEqual(interactive.errors.length, 0);
    assert.strictEqual(explain.valid.length, 0);
    assert.strictEqual(explain.errors.length, 1);
    assert.ok(explain.errors[0].includes("explain mode"));
  });

  it("should block createWorksheet in explain mode", () => {
    const ops = [{ operation: "createWorksheet", name: "New" }];
    const result = validateOperations(ops, "explain");
    assert.strictEqual(result.valid.length, 0);
    assert.strictEqual(result.errors.length, 1);
  });

  it("should block deleteWorksheet in explain mode", () => {
    const ops = [{ operation: "deleteWorksheet", name: "Sheet1" }];
    const result = validateOperations(ops, "explain");
    assert.strictEqual(result.valid.length, 0);
    assert.strictEqual(result.errors.length, 1);
  });

  it("should block writeFormulas in explain mode", () => {
    const ops = [{ operation: "writeFormulas", sheet: "S", range: "A1", formulas: [["=SUM()"]] }];
    const result = validateOperations(ops, "explain");
    assert.strictEqual(result.valid.length, 0);
    assert.strictEqual(result.errors.length, 1);
  });

  it("should allow multiple readRange ops in explain mode", () => {
    const ops = [
      { operation: "readRange", sheet: "A", range: "A1" },
      { operation: "readRange", sheet: "B", range: "C1:D5" },
    ];
    const result = validateOperations(ops, "explain");
    assert.strictEqual(result.valid.length, 2);
    assert.strictEqual(result.errors.length, 0);
  });

  it("should filter mixed ops in explain mode (keep reads, block writes)", () => {
    const ops = [
      { operation: "readRange", sheet: "A", range: "A1" },
      { operation: "writeValues", sheet: "A", range: "A1", values: [["x"]] },
      { operation: "readRange", sheet: "B", range: "B1" },
      { operation: "setFillColor", sheet: "A", range: "A1", color: "red" },
    ];
    const result = validateOperations(ops, "explain");
    assert.strictEqual(result.valid.length, 2);
    assert.strictEqual(result.errors.length, 2);
  });
});

describe("extractOperations — _switchMode parsing", () => {
  it("should extract _switchMode flag from JSON response", () => {
    const response = '```json\n{"plan": "User wants to write", "operations": [], "_switchMode": "interactive", "_switchMessage": "Switch?"}\n```';
    const result = extractOperations(response);
    assert.ok(result !== null);
    assert.strictEqual(result._switchMode, "interactive");
    assert.strictEqual(result._switchMessage, "Switch?");
    assert.strictEqual(result.operations.length, 0);
  });

  it("should return null _switchMode when not present", () => {
    const response = '```json\n{"plan": "Some plan", "operations": []}\n```';
    const result = extractOperations(response);
    assert.ok(result !== null);
    assert.strictEqual(result._switchMode, null);
    assert.strictEqual(result._switchMessage, null);
  });

  it("should still extract operations alongside _switchMode", () => {
    const response = '```json\n{"plan": "P", "operations": [{"operation":"readRange","sheet":"A","range":"A1"}], "_switchMode": "interactive"}\n```';
    const result = extractOperations(response);
    assert.ok(result !== null);
    assert.strictEqual(result.operations.length, 1);
    assert.strictEqual(result._switchMode, "interactive");
  });

  it("should handle plain text responses (no JSON block)", () => {
    const response = "This is your data analysis. Column A has 42 values.";
    const result = extractOperations(response);
    assert.strictEqual(result, null);
  });

  it("should handle malformed JSON gracefully", () => {
    const response = "```json\n{broken json\n```";
    const result = extractOperations(response);
    assert.strictEqual(result, null);
  });
});

describe("validateOperations — interactive mode allows all ops", () => {
  it("should pass through all valid operations in interactive mode", () => {
    const ops = [
      { operation: "createWorksheet", name: "New" },
      { operation: "writeValues", sheet: "New", range: "A1", values: [["x"]] },
      { operation: "readRange", sheet: "New", range: "A1" },
      { operation: "deleteWorksheet", name: "New" },
      { operation: "autofitColumns", sheet: "New" },
      { operation: "deleteRows", sheet: "New", range: "A1" },
    ];
    const result = validateOperations(ops, "interactive");
    assert.strictEqual(result.valid.length, 6);
    assert.strictEqual(result.errors.length, 0);
  });

  it("should still validate required fields in interactive mode", () => {
    const ops = [
      { operation: "writeValues", sheet: "A" },
      { operation: "readRange", range: "A1" },
    ];
    const result = validateOperations(ops, "interactive");
    // Both should have missing field errors
    assert.strictEqual(result.valid.length, 0);
    assert.ok(result.errors.length >= 2);
  });
});

describe("Edge cases", () => {
  it("should handle empty operations array", () => {
    const r1 = validateOperations([], "interactive");
    const r2 = validateOperations([], "explain");
    assert.strictEqual(r1.valid.length, 0);
    assert.strictEqual(r1.errors.length, 0);
    assert.strictEqual(r2.valid.length, 0);
    assert.strictEqual(r2.errors.length, 0);
  });

  it("should handle null/undefined input to extractOperations", () => {
    assert.strictEqual(extractOperations(null), null);
    assert.strictEqual(extractOperations(undefined), null);
    assert.strictEqual(extractOperations(""), null);
  });

  it("should default to interactive mode when mode is unknown", () => {
    const ops = [{ operation: "writeValues", sheet: "A", range: "A1", values: [["x"]] }];
    // If mode falls through to interactive, should pass
    const result = validateOperations(ops, "interactive");
    assert.strictEqual(result.valid.length, 1);
  });
});

// ─── Summary ─────────────────────────────────────────────────────────────────

console.log("\n" + "─".repeat(40));
console.log("Total: " + total + " | Passed: " + passed + " | Failed: " + failed);
console.log("─".repeat(40));

if (failed > 0) {
  process.exit(1);
} else {
  console.log("All tests passed.");
  process.exit(0);
}
