/* global console, document */

/* Thin orchestrator — wires together store, executor, and UI init */

import {
  S,
  fullReset,
  resetForNextTask as storeResetTask,
  stopAgent as storeStopAgent,
} from "./store.js";
import { abortActiveRequest } from "./agent/ai.js";
import { appendMessage, showWelcome, resetStepGroup } from "./ui/chat.js";
import { switchMode as switchModeUI, clearChatUI } from "./ui/config.js";
import { init as initTelemetry, flush as flushTelemetry } from "./telemetry.js";
import { initAgentChat as initAgentChatUI } from "./ui/init.js";
import { handleImprove as executorHandleImprove } from "./agent/executor.js";

// ─── Public API — initAgentChat ──────────────────────────────────────────────

/**
 * Initialize the agent chat UI.
 * Wires up callbacks for acknowledge and improve flows.
 */
export async function initAgentChat() {
  await initAgentChatUI(acknowledgeCompletion, handleImprove);
}

// ─── Acknowledge / Completion ────────────────────────────────────────────────

function acknowledgeCompletion() {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return;

  const feedbackGroup = messagesContainer.querySelector(".agent-feedback-group");
  if (feedbackGroup) feedbackGroup.style.display = "none";

  const improveArea = messagesContainer.querySelector(".agent-improve-area");
  if (improveArea) improveArea.style.display = "none";

  S.isExecuting = false;
  S.isStopped = false;
}

// ─── Improve Flow ────────────────────────────────────────────────────────────

async function handleImprove(comment, taskStats) {
  await executorHandleImprove(comment, taskStats, acknowledgeCompletion);
}

// ─── Stop Handler ────────────────────────────────────────────────────────────

export function stopAgent() {
  storeStopAgent();
}

// ─── Mode Switching ──────────────────────────────────────────────────────────

export function switchMode(mode) {
  const prev = S.currentMode;
  const next = switchModeUI(mode, appendMessage) || prev;
  S.currentMode = next;
}

// ─── Clear / Reset ───────────────────────────────────────────────────────────

export function clearChat() {
  // Kill any running AI request immediately
  fullReset();
  abortActiveRequest();

  flushTelemetry(false);
  clearChatUI(showWelcome);
  resetStepGroup();

  console.log("[agentChat] Chat cleared");
}

export function resetForNextTask() {
  storeResetTask();

  flushTelemetry(true);
  initTelemetry();
  clearChatUI(showWelcome);
  resetStepGroup();

  console.log("[agentChat] Reset for next task");
}
