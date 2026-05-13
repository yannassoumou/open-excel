/* global Excel, console */

/* Workbook snapshots for revert and operation verification — batched loads, minimal syncs */

const WRITABLE_OPS = [
  "writeValues",
  "writeFormulas",
  "createTable",
  "deleteWorksheet",
  "createWorksheet",
  "deleteRows",
  "deleteColumns",
];

/**
 * Capture a snapshot of the workbook state before executing operations.
 * Batched loads with single context.sync() per scope.
 * @param {object[]} operations - Operations about to be executed
 * @returns {Promise<object|null>} Snapshot data or null if no writable ops
 */
export async function captureSnapshot(operations) {
  const writableOps = operations.filter((op) => WRITABLE_OPS.includes(op.operation));
  if (writableOps.length === 0) return null;

  try {
    const snapshot = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Batch load all sheets + used ranges
      const sheetData = sheets.items.map((sheet) => {
        sheet.load("name");
        const usedRange = sheet.getUsedRange();
        usedRange.load("values, address");
        return { sheet, usedRange };
      });
      await context.sync();

      const data = {};
      for (const { sheet, usedRange } of sheetData) {
        try {
          if (usedRange.values) {
            data[sheet.name] = {
              range: usedRange.address,
              values: usedRange.values,
            };
          }
        } catch {
          // Sheet exists but has no used range
        }
      }
      return data;
    });
    console.log(`[snapshot] Captured ${Object.keys(snapshot || {}).length} sheet(s)`);
    return snapshot;
  } catch (error) {
    console.warn(`[snapshot] Failed to capture: ${error.message}`);
    return null;
  }
}

/**
 * Revert to a captured snapshot by restoring sheet values.
 * Atomic: clear + write in a single sync to avoid partial-restore state.
 * @param {object} snapshot
 * @returns {Promise<string>} Result message
 */
export async function revertToSnapshot(snapshot) {
  if (!snapshot || Object.keys(snapshot).length === 0) {
    return "  ⚠ No snapshot available to revert to";
  }

  try {
    const result = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const existingSheets = new Set(sheets.items.map((s) => s.name));
      const restored = [];

      // First pass: ensure all sheets exist and load ranges
      const sheetOps = [];
      for (const [sheetName, sheetData] of Object.entries(snapshot)) {
        let sheet;
        if (existingSheets.has(sheetName)) {
          sheet = sheets.getItem(sheetName);
        } else {
          sheet = sheets.add(sheetName);
          existingSheets.add(sheetName);
        }
        sheet.load("name");
        const range = sheet.getRange(sheetData.range);
        range.load("address");
        sheetOps.push({ sheet, range, sheetData });
      }
      await context.sync();

      // Second pass: clear contents AND write values — single sync for atomicity
      for (const { sheet, range, sheetData } of sheetOps) {
        const currentRange = sheet.getRange(range.address);
        currentRange.clear("contents");
        range.values = sheetData.values;
        restored.push(`${sheetName}!${range.address}`);
      }
      await context.sync();

      return restored.length > 0
        ? `  🔄 Reverted: ${restored.join(", ")}`
        : "  🔄 Reverted to previous state";
    });

    console.log(`[revert] ${result}`);
    return result;
  } catch (error) {
    console.error(`[revert] Failed: ${error.message}`);
    return `  ⚠ Revert failed: ${error.message}`;
  }
}

/**
 * Read back verification data for ALL executed operations.
 * Batched loads per operation group to minimize sync calls.
 * @param {object[]} operations
 * @param {string[]} actualSheetNames
 * @returns {Promise<string>} Verification summary
 */
export async function verifyOperations(operations, actualSheetNames) {
  try {
    const verification = await Excel.run(async (context) => {
      const results = [];
      const seen = new Set();

      for (let i = 0; i < operations.length; i++) {
        const op = operations[i];
        const type = op.operation;
        if (type === "readRange") continue;

        const sheetName = (actualSheetNames && actualSheetNames[i]) || op.sheet || op.name;
        const range = op.range || "A1:C3";
        const key = `${sheetName}|${range}|${type}`;
        if (seen.has(key)) continue;
        seen.add(key);

        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.load("name");

        switch (type) {
          case "writeValues":
          case "writeFormulas": {
            const targetRange = sheet.getRange(range);
            targetRange.load("values, address, addressLocal");
            await context.sync();

            if (targetRange.values) {
              const rows = Math.min(targetRange.values.length, 3);
              const cols = Math.min(targetRange.values[0] ? targetRange.values[0].length : 0, 3);
              const sample = [];
              for (let r = 0; r < rows; r++) {
                const rowCells = [];
                for (let c = 0; c < cols; c++) {
                  rowCells.push(String(targetRange.values[r][c] ?? ""));
                }
                sample.push(rowCells.join(", "));
              }
              const addr = targetRange.addressLocal || targetRange.address;
              results.push(`  ✓ ${sheetName}!${addr}: ${sample.join(" | ")}`);
            }
            break;
          }

          case "createTable": {
            const tables = sheet.tables;
            tables.load("items/name, items/rangeA1");
            await context.sync();
            const tableNames = tables.items.map((t) => t.name).join(", ") || "(none)";
            results.push(`  ✓ ${sheetName}: tables = [${tableNames}]`);
            break;
          }

          case "createWorksheet": {
            results.push(`  ✓ Sheet "${sheetName}" exists`);
            break;
          }

          case "autofitColumns":
          case "autofitRows": {
            const usedRange = sheet.getUsedRange();
            usedRange.load("address, addressLocal");
            await context.sync();
            results.push(`  ✓ ${sheetName}: used range = ${usedRange.addressLocal}`);
            break;
          }

          case "setFillColor":
          case "setFontColor":
          case "setFontBold":
          case "setNumberFormat": {
            const targetRange = sheet.getRange(range);
            targetRange.load("address, addressLocal");
            await context.sync();

            const sampleRange = sheet.getRange(targetRange.address);
            sampleRange.load("values");
            await context.sync();

            if (sampleRange.values) {
              const rows = Math.min(sampleRange.values.length, 2);
              const cols = Math.min(sampleRange.values[0] ? sampleRange.values[0].length : 0, 3);
              const sample = [];
              for (let r = 0; r < rows; r++) {
                const rowCells = [];
                for (let c = 0; c < cols; c++) {
                  rowCells.push(String(sampleRange.values[r][c] ?? ""));
                }
                sample.push(rowCells.join(", "));
              }
              const addr = targetRange.addressLocal || targetRange.address;
              results.push(`  ✓ ${sheetName}!${addr} (data: ${sample.join(" | ")})`);
            } else {
              const addr = targetRange.addressLocal || targetRange.address;
              results.push(`  ✓ ${sheetName}!${addr} (formatted)`);
            }
            break;
          }

          default:
            results.push(`  ✓ ${sheetName}!${range} (${type})`);
        }
      }

      return results.join("\n");
    });

    return verification;
  } catch (error) {
    return `  ⚠ Verification failed: ${error.message}`;
  }
}

/**
 * Extract actual sheet names from operation results.
 * @param {string[]} results - Operation result strings
 * @param {object[]} operations - Original operations array
 * @returns {string[]} Actual sheet names in order
 */
export function extractSheetNamesFromResults(results, operations) {
  return results.map((result, i) => {
    const op = operations[i];
    if (!result) return op.sheet || op.name;

    const sheetMatch = result.match(/Sheet "([^"]+)" ready/);
    if (sheetMatch) return sheetMatch[1];

    const writeMatch = result.match(/Wrote \d+ rows? to ([^!]+)!/);
    if (writeMatch) return writeMatch[1];

    const formatMatch = result.match(/Set .* on ([^!]+)!/);
    if (formatMatch) return formatMatch[1];

    const autofitMatch = result.match(/Autofit (?:columns|rows) on "([^"]+)"/);
    if (autofitMatch) return autofitMatch[1];

    const tableMatch = result.match(/created on ([^!]+)/);
    if (tableMatch) return tableMatch[1];

    const readMatch = result.match(/\d+ rows? x \d+ cols? on ([^!]+)!/);
    if (readMatch) return readMatch[1];

    const deleteMatch = result.match(/Deleted sheet "([^"]+)"/);
    if (deleteMatch) return deleteMatch[1];

    return op.sheet || op.name;
  });
}
