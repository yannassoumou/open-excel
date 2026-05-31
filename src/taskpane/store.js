/* Central state store for the agent chat system.
 * Single source of truth for all mutable state. */

// ─── Constants ───────────────────────────────────────────────────────────────

export const MAX_CONSECUTIVE_FAILURES = 3;
export const MAX_NO_OPS_RETRIES = 2;

// ─── State ───────────────────────────────────────────────────────────────────

// Exported for internal use by agentChat modules.
// Prefer using the getter/setter/action APIs below when possible.
export const S = {
  // Host & mode
  currentHost: localStorage.getItem("agentHost") || "excel",
  currentMode: localStorage.getItem("agentMode") || "interactive",

  // Execution state
  isExecuting: false,
  isStopped: false,

  // Conversation
  conversationHistory: [],
  stepStack: [],
  originalQuery: "",
  pendingImageBase64: null,

  // Retry tracking
  consecutiveFailures: 0,
  noOpsRetries: 0,

  // Feedback
  feedbackStepData: null,
  feedbackTotalSteps: 0,
  feedbackTotalTiming: 0,
};

// ─── Getters ─────────────────────────────────────────────────────────────────

export function getState() {
  return S;
}

export function getHost() {
  return S.currentHost;
}

export function getMode() {
  return S.currentMode;
}

export function isExecuting() {
  return S.isExecuting;
}

export function isStopped() {
  return S.isStopped;
}

export function getConversationHistory() {
  return S.conversationHistory;
}

export function getStepStack() {
  return S.stepStack;
}

// ─── Setters ─────────────────────────────────────────────────────────────────

export function setHost(host) {
  S.currentHost = host;
  localStorage.setItem("agentHost", host);
}

export function setMode(mode) {
  S.currentMode = mode;
  localStorage.setItem("agentMode", mode);
}

export function setExecuting(val) {
  S.isExecuting = val;
}

export function setStopped(val) {
  S.isStopped = val;
}

// ─── Actions ─────────────────────────────────────────────────────────────────

/**
 * Push a message to conversation history.
 */
export function pushMessage(role, content) {
  S.conversationHistory.push({ role, content });
}

/**
 * Push a step to the step stack.
 */
export function pushStep(step) {
  S.stepStack.push(step);
}

/**
 * Reset state for a new task (preserves host + mode).
 */
export function resetForNextTask() {
  S.conversationHistory = [];
  S.stepStack = [];
  S.originalQuery = "";
  S.feedbackStepData = null;
  S.feedbackTotalSteps = 0;
  S.feedbackTotalTiming = 0;
  S.consecutiveFailures = 0;
  S.noOpsRetries = 0;
}

/**
 * Full reset (clear chat, stop execution).
 */
export function fullReset() {
  S.isStopped = true;
  S.isExecuting = false;
  S.conversationHistory = [];
  S.stepStack = [];
  S.originalQuery = "";
  S.pendingImageBase64 = null;
  S.feedbackStepData = null;
  S.feedbackTotalSteps = 0;
  S.feedbackTotalTiming = 0;
  S.consecutiveFailures = 0;
  S.noOpsRetries = 0;
}

/**
 * Stop agent execution.
 */
export function stopAgent() {
  S.isStopped = true;
  S.isExecuting = false;
  console.log("[store] Agent stopped by user");
}

/**
 * Increment consecutive failures counter.
 */
export function incrementFailures() {
  S.consecutiveFailures++;
}

/**
 * Reset consecutive failures counter.
 */
export function resetFailures() {
  S.consecutiveFailures = 0;
}

/**
 * Set the original user query.
 */
export function setOriginalQuery(query) {
  S.originalQuery = query;
}

/**
 * Set pending image base64 data.
 */
export function setPendingImage(base64) {
  S.pendingImageBase64 = base64;
}

/**
 * Set feedback step data.
 */
export function setFeedbackStepData(data) {
  S.feedbackStepData = data;
}

/**
 * Increment feedback step counter and timing.
 */
export function addFeedbackStep(timingMs) {
  S.feedbackTotalSteps++;
  S.feedbackTotalTiming += timingMs;
}

/**
 * Set feedback task stats.
 */
export function getTaskStats() {
  return {
    totalSteps: S.feedbackTotalSteps,
    totalTiming: S.feedbackTotalTiming,
    finalSuccess: true,
  };
}
