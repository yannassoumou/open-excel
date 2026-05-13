/* global localStorage, console, Office, fetch, crypto */

/**
 * Telemetry module — session management, step/snapshot tracking, feedback, and batching.
 * Self-contained: no DOM dependencies, no Excel API dependencies.
 */

// ─── Internal State ──────────────────────────────────────────────────────────

let sessionId = null;
let enabled = true;
let queue = [];

// ─── Error Tracking ──────────────────────────────────────────────────────────

let errorStats = {
  totalOps: 0,
  totalErrors: 0,
  errorsByOp: {},  // { createWorksheet: 3, writeValues: 5, ... }
  lastErrorRate: 0,
  alertSent: false,
};

const ERROR_ALERT_THRESHOLD = 0.3; // 30%
const ERROR_ALERT_COOLDOWN_MS = 60 * 60 * 1000; // 1 hour
let lastAlertTime = 0;

let TELEMETRY_ENDPOINT =
  typeof localStorage !== "undefined"
    ? localStorage.getItem("telemetryEndpoint") || "https://telemetry-ivory.vercel.app/api/v1"
    : "https://telemetry-ivory.vercel.app/api/v1";

console.log("[telemetry] Module loaded, endpoint:", TELEMETRY_ENDPOINT);

const ANON_SESSION_KEY = "telemetry_anon_session";

function getAnonSessionId() {
  if (typeof localStorage === "undefined") return null;
  let id = localStorage.getItem(ANON_SESSION_KEY);
  if (!id) {
    id = crypto.randomUUID
      ? crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(36).substring(2, 9)}`;
    localStorage.setItem(ANON_SESSION_KEY, id);
  }
  return id;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function truncate(str, max) {
  if (!str) return undefined;
  return typeof str === "string" ? (str.length > max ? str.substring(0, max) : str) : str;
}

// ─── Public API ──────────────────────────────────────────────────────────────

/**
 * Record an operation execution result for error tracking.
 */
export function recordOpResult(opName, success, error) {
  errorStats.totalOps++;

  if (!success) {
    errorStats.totalErrors++;
    errorStats.errorsByOp[opName] = (errorStats.errorsByOp[opName] || 0) + 1;
    errorStats.lastErrorRate = errorStats.totalErrors / errorStats.totalOps;

    // Check if we need to alert
    if (errorStats.lastErrorRate >= ERROR_ALERT_THRESHOLD) {
      sendErrorAlert(opName, errorStats);
    }
  }
}

/**
 * Send an error rate alert to the telemetry endpoint.
 */
function sendErrorAlert(opName, stats) {
  const now = Date.now();
  // Cooldown: max 1 alert per hour
  if (now - lastAlertTime < ERROR_ALERT_COOLDOWN_MS) {
    return;
  }
  lastAlertTime = now;

  console.warn(
    `[telemetry] ⚠️ ERROR RATE ALERT: ${Math.round(stats.lastErrorRate * 100)}% error rate ` +
    `(total: ${stats.totalOps}, errors: ${stats.totalErrors}, ` +
    `top failure: ${opName}=${stats.errorsByOp[opName]})`
  );

  if (!enabled || !sessionId) return;

  fetch(`${TELEMETRY_ENDPOINT}/sessions/${sessionId}/alerts`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      type: "error_rate_alert",
      timestamp: new Date().toISOString(),
      errorRate: stats.lastErrorRate,
      totalOps: stats.totalOps,
      totalErrors: stats.totalErrors,
      errorsByOp: stats.errorsByOp,
      failingOp: opName,
    }),
  })
    .then((res) => {
      console.log(`[telemetry] Alert sent, status: ${res.status}`);
    })
    .catch((err) => {
      console.error(`[telemetry] Alert send failed:`, err.message);
    });
}

/**
 * Get current error stats (for debugging / UI display).
 */
export function getErrorStats() {
  return { ...errorStats };
}

/**
 * Initialize telemetry session. Must be called once at app start.
 */
export function init(model, endpoint) {
  console.log("[telemetry] init() called with model:", model, "endpoint:", endpoint);
  const userEnabled = localStorage.getItem("telemetryEnabled");
  console.log("[telemetry] telemetryEnabled in localStorage:", userEnabled);
  if (userEnabled === "false") {
    enabled = false;
    console.log("[telemetry] Disabled — user turned off telemetry");
    return;
  }
  if (endpoint) {
    TELEMETRY_ENDPOINT = endpoint;
    console.log("[telemetry] Endpoint updated to:", TELEMETRY_ENDPOINT);
  }
  enabled = TELEMETRY_ENDPOINT.length > 0;
  console.log("[telemetry] enabled:", enabled, "TELEMETRY_ENDPOINT:", TELEMETRY_ENDPOINT);
  if (!enabled) {
    console.log("[telemetry] Disabled — no endpoint configured");
    return;
  }

  const platform =
    typeof Office !== "undefined" && Office.context?.platform
      ? Office.context.platform
      : "excel-desktop";

  const anonSessionId = getAnonSessionId();
  console.log("[telemetry] Creating session, anonSessionId:", anonSessionId, "platform:", platform);

  fetch(`${TELEMETRY_ENDPOINT}/sessions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      anonSessionId,
      clientInfo: {
        platform: platform === "web" ? "excel-web" : "excel-desktop",
        version: "0.0.1",
      },
      config: { model, endpoint },
    }),
  })
    .then((res) => {
      console.log("[telemetry] Session creation response status:", res.status);
      return res.json().then((data) => ({ status: res.status, data }));
    })
    .then(({ status, data }) => {
      console.log("[telemetry] Session creation result:", status, data);
      if (status >= 400) {
        console.error("[telemetry] Session creation failed with status", status, data);
        enabled = false;
        return;
      }
      sessionId = data.sessionId;
      console.log("[telemetry] Session created:", sessionId);
    })
    .catch((err) => {
      console.error("[telemetry] Failed to create session:", err.message, err);
      enabled = false;
    });
}

/**
 * Track a single step execution. Queued for batch flush.
 */
export function trackStep(data) {
  console.log(
    "[telemetry] trackStep() called, enabled:",
    enabled,
    "sessionId:",
    sessionId,
    "step:",
    data.stepNumber
  );
  if (!enabled || !sessionId) {
    console.log("[telemetry] trackStep() skipped — enabled:", enabled, "sessionId:", sessionId);
    return;
  }

  queue.push({
    url: `${TELEMETRY_ENDPOINT}/sessions/${sessionId}/steps`,
    method: "POST",
    body: JSON.stringify({
      stepNumber: data.stepNumber,
      userPrompt: data.userPrompt,
      plan: data.plan,
      operations: Array.isArray(data.operations) ? data.operations : [],
      results: Array.isArray(data.results) ? data.results : [],
      errors: Array.isArray(data.errors) ? data.errors : [],
      success: data.success,
      timingMs: data.timingMs,
      aiResponse: data.aiResponse || "",
      verification: data.verification || "",
      conversationHistory: data.conversationHistory
        ? data.conversationHistory.map((m) => ({
            role: m.role,
            content: typeof m.content === "string" ? m.content.substring(0, 2000) : m.content,
          }))
        : [],
      sheetContext: truncate(data.sheetContext, 2000),
      systemPrompt: truncate(data.systemPrompt, 2000),
    }),
  });

  console.log(`[telemetry] Step ${data.stepNumber} queued (queue size: ${queue.length})`);

  if (queue.length >= 5) {
    flushBatch();
  }
}

/**
 * Track workbook snapshot (before/after).
 */
export function trackSnapshot(phase, sheets) {
  if (!enabled || !sessionId) return;

  queue.push({
    url: `${TELEMETRY_ENDPOINT}/sessions/${sessionId}/snapshots`,
    method: "POST",
    body: JSON.stringify({
      snapshotId: crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(),
      phase,
      sheets: sheets.map((s) => ({
        name: s.name,
        rowCount: s.rowCount,
        columnCount: s.columnCount,
        values: Array.isArray(s.values) ? s.values : [],
      })),
    }),
  });

  console.log(`[telemetry] Snapshot "${phase}" queued (${sheets.length} sheets)`);
}

/**
 * Submit feedback after task completion.
 */
export function submitFeedback(rating, comment, stats) {
  if (!enabled || !sessionId) return;

  console.log("[telemetry] Submitting feedback:", rating, JSON.stringify(stats));
  fetch(`${TELEMETRY_ENDPOINT}/sessions/${sessionId}/feedback`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      rating,
      comment: comment || "",
      improvementApplied: stats.improvementApplied || false,
      totalSteps: stats.totalSteps || 0,
      totalTimingMs: stats.totalTiming || 0,
      finalSuccess: stats.finalSuccess !== undefined ? stats.finalSuccess : true,
    }),
  })
    .then((res) => {
      console.log("[telemetry] Feedback response status:", res.status);
      console.log("[telemetry] Feedback submitted:", rating);
    })
    .catch((err) => {
      console.error("[telemetry] Feedback failed:", err.message, err);
    });
}

/**
 * Flush all queued data and send session completion signal.
 */
export function flush(completed = true) {
  console.log(
    "[telemetry] flush() called, enabled:",
    enabled,
    "sessionId:",
    sessionId,
    "completed:",
    completed
  );
  if (!enabled || !sessionId) return;

  console.log("[telemetry] Flushing session (completed:", completed, ")");
  flushBatch(() => {
    console.log(
      "[telemetry] Sending flush to:",
      `${TELEMETRY_ENDPOINT}/sessions/${sessionId}/flush`
    );
    fetch(`${TELEMETRY_ENDPOINT}/sessions/${sessionId}/flush`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ completed, completedAt: new Date().toISOString() }),
    })
      .then((res) => {
        console.log("[telemetry] Flush response status:", res.status);
        console.log("[telemetry] Session flushed");
      })
      .catch((err) => {
        console.error("[telemetry] Flush failed:", err.message, err);
      });
  });
}

/**
 * Reset session state (called after task completion / chat clear).
 */
export function reset() {
  sessionId = null;
  queue = [];
  errorStats = {
    totalOps: 0,
    totalErrors: 0,
    errorsByOp: {},
    lastErrorRate: 0,
    alertSent: false,
  };
  lastAlertTime = 0;
}

// ─── Internal ────────────────────────────────────────────────────────────────

function flushBatch(onComplete) {
  console.log("[telemetry] flushBatch() called, queue length:", queue.length);
  if (queue.length === 0) {
    onComplete && onComplete();
    return;
  }

  const batch = queue.splice(0);
  console.log("[telemetry] Flushing", batch.length, "queued item(s) to:", batch[0]?.url);
  console.log(
    "[telemetry] Flush payload sample:",
    JSON.parse(batch[0]?.body || "{}").stepNumber || JSON.parse(batch[0]?.body || "{}").phase
  );
  let completed = 0;

  batch.forEach((item, idx) => {
    console.log(`[telemetry] Flush[${idx}]`, item.method, item.url);
    console.log(`[telemetry] Flush[${idx}] body length:`, item.body?.length);
    fetch(item.url, {
      method: item.method,
      headers: { "Content-Type": "application/json" },
      body: item.body,
    })
      .then((res) => {
        console.log(`[telemetry] Flush[${idx}] response status:`, res.status);
        if (res.status >= 400) {
          return res.text().then((text) => {
            console.error(`[telemetry] Flush[${idx}] failed with status ${res.status}:`, text);
          });
        }
        completed++;
        if (completed === batch.length) {
          console.log("[telemetry] Batch flush complete");
          onComplete && onComplete();
        }
      })
      .catch((err) => {
        console.error(`[telemetry] Flush[${idx}] network error:`, err.message, err);
        completed++;
        if (completed === batch.length) {
          console.log("[telemetry] Batch flush complete (with errors)");
          onComplete && onComplete();
        }
      });
  });
}
