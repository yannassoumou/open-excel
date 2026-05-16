/* global Excel, console */

/* Excel workbook context reading helpers — batched loads, single sync per scope */

/**
 * Read current workbook context (all sheets + sample data).
 * Uses batched load with single context.sync() — Microsoft recommended pattern.
 */
export async function getSheetContext() {
  try {
    const info = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const result = {
        sheetNames: sheets.items.map((s) => s.name),
        sheets: [],
      };

      // Batch load all used ranges
      const sheetRanges = sheets.items.map((sheet) => {
        sheet.load("name");
        const usedRange = sheet.getUsedRange(false);
        usedRange.load("address, rowCount, columnCount, values");
        return { sheet, usedRange };
      });
      await context.sync();

      for (const { sheet, usedRange } of sheetRanges) {
        result.sheets.push({
          name: sheet.name,
          address: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          values: usedRange.values,
        });
      }

      return result;
    });

    console.log(`[context] Read ${info.sheets.length} sheet(s):`, info.sheetNames);

    const lines = [];
    for (const sheet of info.sheets) {
      lines.push(
        `Sheet "${sheet.name}": ${sheet.address} (${sheet.rowCount} rows x ${sheet.columnCount} cols)`
      );

      if (sheet.values && sheet.values.length > 0) {
        const normalized = sheet.values.map((row) =>
          row.map((cell) => (cell === undefined || cell === null ? "" : cell))
        );
        const totalRows = normalized.length;
        const totalCols = normalized[0] ? normalized[0].length : 0;
        const rows = Math.min(totalRows, 100);
        const cols = Math.min(totalCols, 30);
        const sample = [];
        for (let r = 0; r < rows; r++) {
          const rowCells = [];
          for (let c = 0; c < cols; c++) {
            const val = normalized[r][c];
            rowCells.push(val === "" ? "" : String(val));
          }
          // Skip fully empty rows
          if (rowCells.some((v) => v !== "")) {
            sample.push(rowCells.join("\t"));
          }
        }
        const truncNote =
          totalRows > rows || totalCols > cols
            ? ` (showing ${rows}/${totalRows} rows, ${cols}/${totalCols} cols)`
            : "";
        lines.push(`  Data${truncNote}:`);
        lines.push(`  ${sample.join("\n  ")}`);
      }
    }

    return `Workbook has ${info.sheetNames.length} sheet(s): ${info.sheetNames.join(
      ", "
    )}\n${lines.join("\n")}\n\nUse sheet names exactly as listed above.`;
  } catch (error) {
    console.warn("[excel:context] Failed to read sheet context:", error);
    return "No sheet context available.";
  }
}

/**
 * Read the FULL workbook context — all sheets, all data.
 * Used for pre-completion validation to ensure nothing was missed.
 */
export async function readFullWorkbookContext() {
  try {
    const info = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const result = {
        sheetNames: sheets.items.map((s) => s.name),
        sheets: [],
      };

      // Batch load all sheets + used ranges in one pass
      const sheetData = sheets.items.map((sheet) => {
        sheet.load("name");
        const usedRange = sheet.getUsedRange();
        usedRange.load("values, rowCount, columnCount");
        return { sheet, usedRange };
      });
      await context.sync();

      for (const { sheet, usedRange } of sheetData) {
        try {
          if (usedRange.rowCount > 0 && usedRange.columnCount > 0) {
            const values = usedRange.values.map((row) =>
              row.map((cell) => (cell === undefined || cell === null ? "" : cell))
            );
            result.sheets.push({
              name: sheet.name,
              rowCount: usedRange.rowCount,
              columnCount: usedRange.columnCount,
              values: values,
            });
          } else {
            result.sheets.push({
              name: sheet.name,
              rowCount: 0,
              columnCount: 0,
              values: [],
            });
          }
        } catch {
          result.sheets.push({
            name: sheet.name,
            rowCount: 0,
            columnCount: 0,
            values: [],
          });
        }
      }

      return result;
    });

    const lines = [];
    lines.push(
      `Final workbook state (${info.sheetNames.length} sheet(s)): ${info.sheetNames.join(", ")}`
    );
    for (const sheet of info.sheets) {
      if (sheet.rowCount > 0) {
        lines.push(`  Sheet "${sheet.name}": ${sheet.rowCount} rows x ${sheet.columnCount} cols`);
        const displayRows = Math.min(sheet.values.length, 10);
        const displayCols = Math.min(sheet.values[0] ? sheet.values[0].length : 0, 15);
        for (let r = 0; r < displayRows; r++) {
          const rowCells = [];
          for (let c = 0; c < displayCols; c++) {
            const val = sheet.values[r][c];
            rowCells.push(val === "" ? "(empty)" : JSON.stringify(val));
          }
          lines.push(`    Row ${r + 1}: ${rowCells.join(" | ")}`);
        }
        if (sheet.values.length > 10) {
          lines.push(`    ... and ${sheet.values.length - 10} more rows`);
        }
      } else {
        lines.push(`  Sheet "${sheet.name}": empty`);
      }
    }
    return lines.join("\n");
  } catch (error) {
    console.warn("[excel:context] Failed to read full workbook context:", error);
    return "Could not read workbook state for validation.";
  }
}

/**
 * Read workbook metadata for telemetry — includes sampled cell values
 * (first 5 rows x 10 cols) so the dashboard shows what the AI saw.
 * Returns array of { name, rowCount, columnCount, values } or null on failure.
 */
export async function readWorkbookMetadata() {
  try {
    const info = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Batch load all sheets + used ranges
      const sheetData = sheets.items.map((sheet) => {
        sheet.load("name");
        const usedRange = sheet.getUsedRange();
        usedRange.load("values, rowCount, columnCount");
        return { sheet, usedRange };
      });
      await context.sync();

      const result = [];
      for (const { sheet, usedRange } of sheetData) {
        try {
          const values = usedRange.values || [];
          const normalized = values.map((row) =>
            row.map((cell) => (cell === undefined || cell === null ? "" : cell))
          );
          const rowCount = usedRange.rowCount;
          const columnCount = usedRange.columnCount;

          const sampleRows = Math.min(normalized.length, 5);
          const sampleCols = Math.min(normalized[0] ? normalized[0].length : 0, 10);
          const sample = [];
          for (let r = 0; r < sampleRows; r++) {
            const rowCells = [];
            for (let c = 0; c < sampleCols; c++) {
              rowCells.push(normalized[r][c]);
            }
            sample.push(rowCells);
          }

          result.push({
            name: sheet.name,
            rowCount,
            columnCount,
            values: sample,
          });
        } catch {
          result.push({ name: sheet.name, rowCount: 0, columnCount: 0, values: [] });
        }
      }
      return result;
    });
    return info;
  } catch (error) {
    console.warn("[excel:context] Failed to read workbook metadata:", error);
    return null;
  }
}

/**
 * Read comprehensive workbook state including tables, charts, pivot tables, shapes, slicers, named items, and conditional formats.
 * Used to give the AI agent full awareness of the workbook structure.
 */
export async function getFullWorkbookState() {
  try {
    const info = await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const result = {
        sheetNames: sheets.items.map((s) => s.name),
        tables: [],
        charts: [],
        pivotTables: [],
        shapes: [],
        slicers: [],
        namedItems: [],
        conditionalFormats: [],
      };

      // Read tables
      try {
        const tables = context.workbook.tables;
        tables.load("items/name, items/rangeA1, items/showHeader, items/showTotals");
        await context.sync();
        result.tables = tables.items.map((t) => ({
          name: t.name,
          range: t.rangeA1,
          showHeader: t.showHeader,
          showTotals: t.showTotals,
        }));
      } catch (tableErr) {
        console.warn("[context] Failed to read tables:", tableErr.message);
      }

      // Read charts, shapes, slicers per sheet
      try {
        const sheetDetails = sheets.items.map((sheet) => {
          sheet.load("name");
          const charts = sheet.charts;
          charts.load("items/name");
          const shapes = sheet.shapes;
          shapes.load("items/name, items/type");
          return { sheet, charts, shapes };
        });
        await context.sync();

        for (const { sheet, charts, shapes } of sheetDetails) {
          try {
            result.charts.push(
              ...charts.items.map((c) => ({
                sheet: sheet.name,
                name: c.name,
              }))
            );
          } catch (e) {
            console.warn("[context] Charts error:", e.message);
          }
          try {
            result.shapes.push(
              ...shapes.items.map((s) => ({
                sheet: sheet.name,
                name: s.name,
                type: s.type,
              }))
            );
          } catch (e) {
            console.warn("[context] Shapes error:", e.message);
          }
        }
      } catch (detailErr) {
        console.warn("[context] Failed to read sheet details:", detailErr.message);
      }

      // Read named items
      try {
        const namedItems = context.workbook.names;
        namedItems.load("items/name, items/value");
        await context.sync();
        result.namedItems = namedItems.items.map((n) => ({
          name: n.name,
          value: n.value,
        }));
      } catch (nameErr) {
        console.warn("[context] Failed to read named items:", nameErr.message);
      }

      return result;
    });

    const lines = [];
    lines.push(`Workbook state: ${info.sheetNames.length} sheet(s): ${info.sheetNames.join(", ")}`);

    if (info.tables.length > 0) {
      lines.push(`Tables (${info.tables.length}):`);
      for (const t of info.tables) {
        lines.push(
          `  - "${t.name}" at ${t.range} (header=${t.showHeader}, totals=${t.showTotals})`
        );
      }
    }

    if (info.charts.length > 0) {
      lines.push(`Charts (${info.charts.length}):`);
      for (const c of info.charts) {
        lines.push(`  - "${c.name}" on ${c.sheet}`);
      }
    }

    if (info.shapes.length > 0) {
      lines.push(`Shapes (${info.shapes.length}):`);
      for (const s of info.shapes) {
        lines.push(`  - "${s.name}" on ${s.sheet} (type=${s.type})`);
      }
    }

    if (info.namedItems.length > 0) {
      lines.push(`Named Items (${info.namedItems.length}):`);
      for (const n of info.namedItems) {
        lines.push(`  - "${n.name}" = "${n.value}"`);
      }
    }

    return lines.join("\n");
  } catch (error) {
    console.warn("[excel:context] Failed to read full workbook state:", error);
    return null;
  }
}
