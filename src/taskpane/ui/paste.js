/* global Excel, document */

/* Paste detection and preview UI */

import { escapeHtml } from "./chat.js";

/**
 * Parse TSV (tab-separated values) from clipboard into a 2D matrix.
 * @param {string} text - Raw pasted text
 * @returns {{ matrix: string[][], rows: number, cols: number } | null}
 */
export function parseTSV(text) {
  if (!text || typeof text !== "string") return null;
  const lines = text.split(/\r?\n/).filter((line) => line.trim() !== "");
  if (lines.length === 0) return null;

  const hasTabs = text.includes("\t");
  const hasSemicolons = text.includes(";");
  const hasCommas = text.includes(",") && lines.length > 1 && lines.every((l) => l.includes(","));

  if (!hasTabs && !hasSemicolons && !hasCommas) return null;

  const delimiter = hasTabs ? "\t" : hasSemicolons ? ";" : ",";
  const matrix = lines.map((line) => line.split(delimiter).map((cell) => cell.trim()));

  const maxCols = Math.max(...matrix.map((row) => row.length));
  for (const row of matrix) {
    while (row.length < maxCols) row.push("");
  }

  return { matrix, rows: matrix.length, cols: maxCols };
}

/**
 * Show a paste preview UI with "Write to Excel" and "Send to AI" buttons.
 */
export function showPastePreview(matrix, rows, cols, callbacks) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return;

  const previewEl = document.createElement("div");
  previewEl.className = "paste-preview";

  const tableHtml = matrix
    .map(
      (row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell || "<empty>")}</td>`).join("")}</tr>`
    )
    .join("");

  previewEl.innerHTML = `
    <div class="paste-preview-header">
      <span>📋 Pasted ${rows} rows × ${cols} cols</span>
      <button class="paste-close" title="Dismiss">✕</button>
    </div>
    <table class="paste-preview-table">
      ${tableHtml}
    </table>
    <div class="paste-preview-actions">
      <button class="paste-write" title="Write this data directly to Excel">📊 Write to Excel</button>
      <button class="paste-send" title="Send as text to the AI agent">💬 Send to AI</button>
    </div>
  `;

  messagesContainer.appendChild(previewEl);
  if (callbacks && callbacks.scrollToBottom) callbacks.scrollToBottom();

  previewEl.querySelector(".paste-close").addEventListener("click", () => {
    if (previewEl.parentElement) previewEl.parentElement.removeChild(previewEl);
  });

  previewEl.querySelector(".paste-write").addEventListener("click", async () => {
    if (previewEl.parentElement) previewEl.parentElement.removeChild(previewEl);

    const result = await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      const usedRange = sheet.getUsedRange();
      usedRange.load("lastRowIndex, lastColumnIndex");
      await context.sync();

      const startRow = Math.max(0, usedRange.lastRowIndex + 1);
      const range = sheet.getRangeByIndexes(startRow, 0, rows, cols);
      range.load("address");
      range.values = matrix;
      await context.sync();

      return `Wrote ${rows}×${cols} data to ${sheet.name}!${range.address}`;
    });

    if (callbacks && callbacks.appendMessage) {
      callbacks.appendMessage("agent", `✅ ${result}`);
    }
    if (callbacks && callbacks.onPasteWritten) {
      callbacks.onPasteWritten(rows, cols);
    }
    if (callbacks && callbacks.scrollToBottom) callbacks.scrollToBottom();
  });

  previewEl.querySelector(".paste-send").addEventListener("click", async () => {
    if (previewEl.parentElement) previewEl.parentElement.removeChild(previewEl);

    // Build a rich prompt for the LLM with context about the pasted data
    const headerRow = matrix[0];
    const dataRows = matrix.slice(1);

    // Detect if first row looks like headers (text-heavy, no pure numbers)
    const headerIsText = headerRow.every((cell) => {
      const trimmed = cell.trim();
      return trimmed === "" || isNaN(Number(trimmed));
    });

    // Build markdown table
    const headerLine = `| ${headerRow.map((c) => c || " ").join(" | ")} |`;
    const separatorLine = `| ${headerRow.map(() => "---").join(" | ")} |`;
    const dataLines = (headerIsText ? dataRows : matrix).map(
      (row) => `| ${row.map((c) => c || " ").join(" | ")} |`
    );
    const table = [headerLine, separatorLine, ...dataLines].join("\n");

    // Detect column types for context hints
    const colHints = headerRow.map((header, ci) => {
      const sample = (headerIsText ? dataRows : matrix).slice(0, 5).map((r) => r[ci]);
      const nums = sample.filter((v) => v.trim() !== "" && !isNaN(Number(v)));
      if (nums.length > sample.length * 0.6) return `${header} (numbers)`;
      return header;
    });

    // Try to get active sheet + cell position for context
    let sheetHint = "";
    try {
      const posInfo = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        const selection = context.workbook.getSelectedRange();
        selection.load("address");
        await context.sync();
        return { sheet: sheet.name, address: selection.address };
      });
      sheetHint = `\nSource: sheet "${posInfo.sheet}", selection ${posInfo.address}`;
    } catch {
      // Ignore — not critical
    }

    const prompt =
      `I just pasted a table from clipboard (${rows} rows × ${cols} cols):${sheetHint}\n\n` +
      table +
      "\n\n" +
      `Columns: ${colHints.join(", ")}` +
      "\n\nWhat should I do with this data?";

    const chatInput = document.getElementById("chat-input");
    const sendButton = document.getElementById("chat-send");
    chatInput.value = prompt;
    chatInput.style.height = "auto";
    chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + "px";
    sendButton.disabled = false;
    chatInput.focus();
  });
}
