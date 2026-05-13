/* global document */

/* Chat UI helpers — message rendering, scrolling, escape */

/**
 * Escape HTML special characters
 */
export function escapeHtml(text) {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

/* ─── Step Group Management ────────────────────────────────────────────────── */

let _activeGroup = null;

/**
 * Create or get the current step group container.
 * All steps are wrapped inside this collapsible group.
 */
function getOrCreateStepGroup() {
  if (_activeGroup && _activeGroup.isConnected) return _activeGroup;

  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return null;

  const group = document.createElement("div");
  group.className = "step-group step-group-collapsed";
  group.innerHTML = `
    <div class="step-group-header" role="button" tabindex="0">
      <span class="step-group-spinner"></span>
      <span class="step-group-label">Working...</span>
      <span class="step-group-toggle">▸</span>
    </div>
    <div class="step-group-body"></div>
  `;

  const header = group.querySelector(".step-group-header");

  // Toggle collapse on click
  header.addEventListener("click", () => {
    const isCollapsed = group.classList.contains("step-group-collapsed");
    group.classList.toggle("step-group-collapsed");
    const toggle = header.querySelector(".step-group-toggle");
    toggle.textContent = isCollapsed ? "▾" : "▸";
  });
  header.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      header.click();
    }
  });

  messagesContainer.appendChild(group);
  _activeGroup = group;
  return group;
}

/**
 * Update the step group header with current progress.
 */
function updateGroupHeader() {
  if (!_activeGroup) return;
  const body = _activeGroup.querySelector(".step-group-body");
  const total = body.querySelectorAll("[data-step]").length;
  const done = body.querySelectorAll(".step-status:not(.step-status-busy)").length;
  const label = _activeGroup.querySelector(".step-group-label");
  const spinner = _activeGroup.querySelector(".step-group-spinner");

  if (total === 0) return;

  const hasErrors = body.querySelectorAll(".status-error").length > 0;
  if (done >= total) {
    // All done
    _activeGroup.classList.remove("step-group-active");
    _activeGroup.classList.add("step-group-done");
    spinner.textContent = hasErrors ? "⚠" : "✓";
    spinner.className = "step-group-icon";
    label.textContent = `${total} step${total > 1 ? "s" : ""} completed`;
    // Auto-collapse when done
    if (!_activeGroup.classList.contains("step-group-collapsed")) {
      setTimeout(() => {
        _activeGroup.classList.add("step-group-collapsed");
        const toggle = _activeGroup.querySelector(".step-group-toggle");
        if (toggle) toggle.textContent = "▸";
      }, 1500);
    }
  } else {
    _activeGroup.classList.add("step-group-active");
    label.textContent = `Step ${done + 1} of ${total}`;
  }
}

/**
 * Finalize the current step group (called when execution ends).
 */
export function finalizeStepGroup() {
  if (!_activeGroup) return;
  updateGroupHeader();
  _activeGroup = null;
}

/**
 * Reset step group state (called on chat clear).
 */
export function resetStepGroup() {
  _activeGroup = null;
}

/**
 * Simple markdown to HTML converter for agent responses.
 * Handles: headers, bold, italic, lists, code blocks, inline code, links.
 */
export function renderMarkdown(text) {
  if (!text) return "";
  let html = escapeHtml(text);

  // Code blocks (```...```)
  html = html.replace(/```(\w*)\n([\s\S]*?)```/g, (_, lang, code) => {
    return `<pre class="md-code-block"><code>${code.trim()}</code></pre>`;
  });

  // Inline code (`code`)
  html = html.replace(/`([^`]+)`/g, '<code class="md-inline-code">$1</code>');

  // Headers
  html = html.replace(/^### (.+)$/gm, "<h4>$1</h4>");
  html = html.replace(/^## (.+)$/gm, "<h3>$1</h3>");
  html = html.replace(/^# (.+)$/gm, "<h2>$1</h2>");

  // Bold and italic
  html = html.replace(/\*\*\*(.+?)\*\*\*/g, "<strong><em>$1</em></strong>");
  html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/\*(.+?)\*/g, "<em>$1</em>");

  // Links
  html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>');

  // Unordered lists (- item or * item)
  html = html.replace(/^[\-\*] (.+)$/gm, "<li>$1</li>");
  html = html.replace(/((?:<li>.*<\/li>\n?)+)/g, "<ul>$1</ul>");

  // Ordered lists (1. item)
  html = html.replace(/^\d+\. (.+)$/gm, "<li>$1</li>");

  // Horizontal rules
  html = html.replace(/^---$/gm, "<hr>");

  // Paragraphs — convert double newlines to <p> tags (skip if already in block element)
  html = html
    .split("\n\n")
    .map((block) => {
      const trimmed = block.trim();
      if (!trimmed) return "";
      if (
        trimmed.startsWith("<h") ||
        trimmed.startsWith("<ul") ||
        trimmed.startsWith("<ol") ||
        trimmed.startsWith("<li") ||
        trimmed.startsWith("<pre") ||
        trimmed.startsWith("<hr")
      ) {
        return trimmed;
      }
      return `<p>${trimmed.replace(/\n/g, "<br>")}</p>`;
    })
    .join("\n");

  return html;
}

/**
 * Scroll chat to bottom only if user is already near the bottom.
 */
export function scrollChatToBottom() {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return;
  const isNearBottom =
    messagesContainer.scrollHeight - messagesContainer.scrollTop - messagesContainer.clientHeight <
    150;
  if (isNearBottom) {
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
  }
}

/**
 * Append a plan message to the chat.
 */
export function appendPlanMessage(planText) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return;

  const planEl = document.createElement("div");
  planEl.className = "agent-message agent";
  planEl.innerHTML = `
    <div class="agent-bubble">
      <div class="plan-header">📋 <strong>Plan</strong></div>
      <div class="plan-body">${escapeHtml(planText)}</div>
    </div>
  `;
  messagesContainer.appendChild(planEl);
  scrollChatToBottom();
}

/**
 * Append a step message showing the operations inside a collapsible block.
 * Steps are added inside a step group container.
 */
export function appendStepMessage(stepNumber, planText, operations) {
  const group = getOrCreateStepGroup();
  if (!group) return;
  const body = group.querySelector(".step-group-body");

  const opText = operations
    .map(
      (op, i) =>
        `${i + 1}. ${op.operation}${op.sheet ? ` → ${op.sheet}` : ""}${op.range ? ` (${op.range})` : ""}`
    )
    .join("\n");

  const codeBlock = opText ? escapeHtml(opText) : "";
  const stepDesc = planText
    ? escapeHtml(planText)
    : escapeHtml(`Executing ${operations.length} operation(s)`);

  const stepEl = document.createElement("div");
  stepEl.className = "agent-message agent";
  stepEl.setAttribute("data-step", stepNumber);
  stepEl.innerHTML = `
    <details class="agent-step-details">
      <summary><span class="step-status step-status-busy"></span> Step ${stepNumber} — ${escapeHtml(operations.length + " op(s)")}</summary>
      <div class="step-content">
        <div class="plan-body">${stepDesc}</div>
        ${codeBlock ? `<pre>${codeBlock}</pre>` : ""}
        <div class="step-result-area"></div>
      </div>
    </details>
  `;
  body.appendChild(stepEl);

  // Auto-expand group when new step arrives
  group.classList.remove("step-group-collapsed");
  group.classList.add("step-group-active");
  const toggle = group.querySelector(".step-group-toggle");
  if (toggle) toggle.textContent = "▾";

  // Auto-collapse previous steps (keep last 2 expanded)
  const allSteps = body.querySelectorAll("[data-step]");
  if (allSteps.length > 2) {
    for (let i = 0; i < allSteps.length - 2; i++) {
      const details = allSteps[i].querySelector("details");
      if (details) details.removeAttribute("open");
    }
  }

  // Expand the new step
  const newDetails = stepEl.querySelector("details");
  if (newDetails) newDetails.setAttribute("open", "");

  updateGroupHeader();
  scrollChatToBottom();
}

/**
 * Append a user or agent message to the chat.
 * @returns {HTMLElement|null}
 */
export function appendMessage(role, text) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return null;

  const messageEl = document.createElement("div");
  messageEl.className = `agent-message ${role}`;

  const bubbleEl = document.createElement("div");
  bubbleEl.className = "agent-bubble";

  if (role === "agent") {
    bubbleEl.innerHTML = renderMarkdown(text);
    bubbleEl.classList.add("md-rendered");
  } else {
    bubbleEl.textContent = text;
  }

  messageEl.appendChild(bubbleEl);
  messagesContainer.appendChild(messageEl);
  scrollChatToBottom();

  return messageEl;
}

/**
 * Append a streaming agent message (typewriter effect).
 * @returns {object} { element, update, complete }
 */
export function appendStreamingMessage() {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return null;

  const messageEl = document.createElement("div");
  messageEl.className = "agent-message agent";

  const bubbleEl = document.createElement("div");
  bubbleEl.className = "agent-bubble md-rendered streaming";
  bubbleEl.innerHTML = "";

  messageEl.appendChild(bubbleEl);
  messagesContainer.appendChild(messageEl);
  scrollChatToBottom();

  let currentText = "";

  return {
    element: messageEl,
    update: (partialText) => {
      currentText = partialText;
      bubbleEl.innerHTML = renderMarkdown(partialText);
      scrollChatToBottom();
    },
    complete: () => {
      bubbleEl.classList.remove("streaming");
      bubbleEl.innerHTML = renderMarkdown(currentText);
    },
  };
}

/**
 * Append execution result to chat with retry/revert buttons.
 * @param {object} result - { success, result/error, stepNumber, operations, hadError }
 * @returns {HTMLElement|null}
 */
export function appendExecutionResult(result) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return null;

  const stepIndex = result.stepNumber !== undefined ? result.stepNumber - 1 : -1;

  const stepEl = messagesContainer.querySelector(
    `.agent-message.agent[data-step="${result.stepNumber}"]`
  );
  let resultArea = null;
  if (stepEl) {
    resultArea = stepEl.querySelector(".step-result-area");
  }

  if (!resultArea) {
    const lastAgentMsg = messagesContainer.querySelector(".agent-message.agent:last-child");
    if (lastAgentMsg) {
      resultArea = lastAgentMsg.querySelector(".step-result-area");
    }
  }

  if (!resultArea) {
    const lastAgentMsg = messagesContainer.querySelector(".agent-message.agent:last-child");
    if (lastAgentMsg) {
      resultArea = lastAgentMsg;
    }
  }

  if (resultArea) {
    const stepElWithResult = resultArea.closest(".agent-message.agent[data-step]");
    if (stepElWithResult) {
      const statusEl = stepElWithResult.querySelector(".step-status");
      if (statusEl) {
        statusEl.classList.remove("step-status-busy");
        statusEl.textContent = result.success ? "✓" : "✗";
        if (!result.success) {
          statusEl.classList.add("status-error");
        }
      }
    }

    const hasSnapshot =
      result.operations &&
      result.operations.some((op) =>
        [
          "writeValues",
          "writeFormulas",
          "createTable",
          "deleteWorksheet",
          "createWorksheet",
        ].includes(op.operation)
      );
    const canRetry = stepIndex >= 0;
    const canRevert = hasSnapshot && stepIndex >= 0;

    let actionButtons = "";
    if (canRetry || canRevert) {
      const buttons = [];
      if (canRetry) {
        buttons.push(
          `<button class="exec-btn exec-btn-retry" data-step="${stepIndex}" title="Retry this step">🔄 Retry</button>`
        );
      }
      if (canRevert) {
        buttons.push(
          `<button class="exec-btn exec-btn-revert" data-step="${stepIndex}" title="Undo this step">↩️ Revert</button>`
        );
      }
      actionButtons = `<div class="exec-actions">${buttons.join("")}</div>`;
    }

    const resultClass = result.success ? "agent-exec-result success" : "agent-exec-result error";
    const icon = result.success ? "✓" : "✗";
    resultArea.innerHTML = `<div class="${resultClass}"><span class="exec-icon">${icon}</span> <span class="exec-text">${escapeHtml(result.result || (result.success ? "Executed successfully" : "Execution failed"))}</span>${actionButtons}</div>`;

    resultArea.querySelectorAll(".exec-btn-retry").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.step, 10);
        result._onRetry && result._onRetry(idx);
      });
    });
    resultArea.querySelectorAll(".exec-btn-revert").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.step, 10);
        result._onRevert && result._onRevert(idx);
      });
    });

    scrollChatToBottom();
    updateGroupHeader();
  }

  return resultArea;
}

/**
 * Show the welcome message in the chat container.
 */
export function showWelcome() {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return;

  messagesContainer.innerHTML = `
    <div class="agent-welcome">
      <h3><img src="assets/icon-80.svg" alt="KuroAgent" style="width:28px;height:28px;vertical-align:middle;margin-right:8px;border-radius:6px;">AI Agent</h3>
      <p>Tell Kuro what you need — it figures out the rest.</p>
      <p style="margin-top: 12px; font-size: 12px; color: #999;">Examples:<br/>
        "Put Hello in A1, make it bold"<br/>
        "Chart the data in A1:B5"<br/>
        "Sum up column A, results in C"</p>
    </div>
  `;
}
