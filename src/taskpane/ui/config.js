/* global document, console, localStorage */

/* Mode switching, chat clear/reset, and feedback UI wiring */

/**
 * Show or hide the agent stop button.
 * @param {boolean} show
 * @param {HTMLElement} stopButton
 * @param {HTMLElement} sendButton
 */
export function setExecutingState(isExecuting, stopButton, sendButton) {
  if (sendButton) sendButton.disabled = isExecuting;
  if (stopButton) stopButton.style.display = isExecuting ? "inline-block" : "none";
}

/**
 * Switch mode and update UI.
 * @param {"interactive"|"explain"} mode
 */
export function switchMode(mode, appendMessage) {
  const validModes = ["interactive", "explain"];
  if (!validModes.includes(mode)) {
    console.error(`[ui:config] Invalid mode: ${mode}`);
    return null;
  }

  localStorage.setItem("agentMode", mode);

  // Update UI toggle buttons
  const interactiveBtn = document.getElementById("mode-interactive");
  const explainBtn = document.getElementById("mode-explain");
  if (interactiveBtn) {
    interactiveBtn.className =
      mode === "interactive" ? "agent-mode-btn mode-active" : "agent-mode-btn";
  }
  if (explainBtn) {
    explainBtn.className = mode === "explain" ? "agent-mode-btn mode-active" : "agent-mode-btn";
  }

  // Update mode badge in header
  const modeBadge = document.getElementById("mode-badge");
  if (modeBadge) {
    modeBadge.textContent = mode === "interactive" ? "⚡ Interactive" : "📖 Explain";
  }

  // Append a system message about the mode switch
  if (appendMessage) {
    const modeLabel = mode === "interactive" ? "interactive" : "explanation";
    appendMessage(
      "agent",
      `Switched to ${modeLabel} mode.` +
        (mode === "explain"
          ? " I'll analyze and explain your Excel data — I cannot make changes."
          : " I can now read and modify your Excel workbook.")
    );
  }

  console.log(`[ui:config] Mode switched to: ${mode}`);
  return mode;
}

/**
 * Clear the chat UI and input.
 */
export function clearChatUI(showWelcome) {
  const messagesContainer = document.getElementById("chat-messages");
  const sendButton = document.getElementById("chat-send");
  const stopButton = document.getElementById("chat-stop");

  if (messagesContainer && showWelcome) {
    showWelcome();
  }

  if (sendButton) sendButton.disabled = true;
  if (stopButton) stopButton.style.display = "none";

  const chatInput = document.getElementById("chat-input");
  if (chatInput) {
    chatInput.value = "";
    chatInput.style.height = "auto";
    chatInput.focus();
  }
}

/**
 * Build the feedback UI element and wire up buttons.
 * @param {HTMLElement} parent - The agent bubble element to append to
 * @param {object} handlers - { onOk, onBad, onImprove }
 */
export function appendFeedbackButtons(parent, handlers) {
  if (!parent) return;

  const feedbackGroup = document.createElement("div");
  feedbackGroup.className = "agent-feedback-group";
  feedbackGroup.innerHTML = `
    <button class="agent-feedback-btn agent-feedback-btn-ok" data-type="ok">👍 OK</button>
    <button class="agent-feedback-btn agent-feedback-btn-improve" data-type="improve">🔧 Improve</button>
    <button class="agent-feedback-btn agent-feedback-btn-bad" data-type="bad">👎 Bad</button>
  `;
  parent.appendChild(feedbackGroup);

  feedbackGroup.querySelector(".agent-feedback-btn-ok").addEventListener("click", () => {
    handlers.onOk && handlers.onOk();
  });

  feedbackGroup.querySelector(".agent-feedback-btn-bad").addEventListener("click", () => {
    handlers.onBad && handlers.onBad();
  });

  feedbackGroup.querySelector(".agent-feedback-btn-improve").addEventListener("click", () => {
    const improveArea = document.createElement("div");
    improveArea.className = "agent-improve-area";
    const uniqueId = Date.now();
    improveArea.innerHTML = `
      <textarea id="improve-comment-${uniqueId}" placeholder="Describe what to improve..." rows="3"></textarea>
      <div class="agent-improve-actions">
        <button class="agent-improve-cancel" id="improve-cancel-${uniqueId}">Cancel</button>
        <button class="agent-improve-submit" id="improve-submit-${uniqueId}">Apply Improvements</button>
      </div>
    `;
    parent.appendChild(improveArea);

    const textarea = improveArea.querySelector(`#improve-comment-${uniqueId}`);

    improveArea.querySelector(`#improve-cancel-${uniqueId}`).addEventListener("click", () => {
      improveArea.remove();
    });

    const submitImprove = () => {
      const comment = textarea.value.trim();
      if (!comment) return;
      improveArea.remove();
      feedbackGroup.style.display = "none";
      handlers.onImprove && handlers.onImprove(comment);
    };

    improveArea
      .querySelector(`#improve-submit-${uniqueId}`)
      .addEventListener("click", submitImprove);

    textarea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        submitImprove();
      }
    });

    textarea.focus();
  });

  return feedbackGroup;
}

/**
 * Build the completion summary message element.
 * @param {object} summaryText - { title, changes, details }
 * @returns {HTMLElement}
 */
export function buildCompletionSummary(summaryText, escapeHtml) {
  const resultEl = document.createElement("div");
  resultEl.className = "agent-message agent";

  const changesHtml =
    summaryText.changes && summaryText.changes.length > 0
      ? `<div class="summary-section"><div class="summary-section-title">Changes made:</div><ul>${summaryText.changes
          .map((c) => `<li>${escapeHtml(c)}</li>`)
          .join("")}</ul></div>`
      : "";

  const detailsHtml = summaryText.details
    ? `<div class="summary-section"><div class="summary-section-title">Details:</div><p>${escapeHtml(summaryText.details)}</p></div>`
    : "";

  resultEl.innerHTML = `
    <div class="agent-bubble">
      <div class="agent-summary-block">
        <div class="summary-header">✅ ${escapeHtml(summaryText.title || "All steps verified and complete!")}</div>
        ${changesHtml}
        ${detailsHtml}
      </div>
    </div>
  `;

  return resultEl;
}
