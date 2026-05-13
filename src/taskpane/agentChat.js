/* global console, document, Excel, PowerPoint */

/* eslint-disable no-undef */

/* Orchestrator — state, init, and main execution loop */

import {
  setConfig as setAiConfig,
  streamFromAI,
  abortActiveRequest,
  getConfig,
  DEFAULTS,
  fetchModels,
  NOUSRESEARCH_ENDPOINT,
} from "./agent/ai.js";
import { buildSystemPrompt } from "./agent/prompts.js";
import {
  extractOperations,
  validateOperations,
  summarizeConversationHistory,
  isCompletionClaim,
  isActionableResponse,
} from "./agent/parser.js";
import { EXCEL_OPERATION_REGISTRY, dispatchExecuteOperation, executeOperationWithTracking } from "./agent/operations.js";
import {
  appendMessage,
  appendPlanMessage,
  appendStepMessage,
  appendExecutionResult,
  showWelcome,
  scrollChatToBottom,
  escapeHtml,
  appendStreamingMessage,
  finalizeStepGroup,
  resetStepGroup,
} from "./ui/chat.js";
import { parseTSV, showPastePreview } from "./ui/paste.js";
import {
  switchMode as switchModeUI,
  clearChatUI,
  appendFeedbackButtons,
  buildCompletionSummary,
} from "./ui/config.js";
import {
  init as initTelemetry,
  trackStep,
  trackSnapshot,
  submitFeedback,
  flush as flushTelemetry,
} from "./telemetry.js";

// ─── Image Preview ───────────────────────────────────────────────────────────

function showImagePreview(base64) {
  const chatContainer = document.getElementById("agent-chat");
  if (!chatContainer) return;
  let previewEl = document.getElementById("image-preview-container");
  if (previewEl) previewEl.remove();

  previewEl = document.createElement("div");
  previewEl.id = "image-preview-container";
  previewEl.className = "image-preview-container";
  previewEl.innerHTML = `
    <div class="image-preview-wrapper">
      <img src="${base64}" class="image-preview-thumb" alt="Attached image" />
      <button class="image-preview-dismiss" title="Remove image">&times;</button>
    </div>
  `;

  const inputArea = chatContainer.querySelector(".agent-input-area");
  if (inputArea) {
    inputArea.insertBefore(previewEl, inputArea.firstChild);
  }

  const dismissBtn = previewEl.querySelector(".image-preview-dismiss");
  if (dismissBtn) {
    dismissBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      pendingImageBase64 = null;
      hideImagePreview();
      imageUploadBtn.style.display = "none";
      chatInput.focus();
    });
  }
}

function hideImagePreview() {
  const previewEl = document.getElementById("image-preview-container");
  if (previewEl) previewEl.remove();
}

/**
 * Append a user message with an image preview to the chat.
 */
function appendUserMessageWithImage(text, imageBase64) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer) return null;

  const messageEl = document.createElement("div");
  messageEl.className = "agent-message user";

  const bubbleEl = document.createElement("div");
  bubbleEl.className = "agent-bubble";

  const textEl = document.createElement("div");
  textEl.className = "user-message-text";
  if (text) {
    textEl.textContent = text;
  }

  const imgEl = document.createElement("img");
  imgEl.className = "user-message-image";
  imgEl.src = imageBase64;
  imgEl.alt = "Attached image";

  bubbleEl.appendChild(textEl);
  bubbleEl.appendChild(imgEl);
  messageEl.appendChild(bubbleEl);
  messagesContainer.appendChild(messageEl);
  scrollChatToBottom();

  return messageEl;
}

// ─── Host Detection ──────────────────────────────────────────────────────────

let currentHost = localStorage.getItem("agentHost") || "excel";
let hostModules = null;

async function loadHostModules() {
  if (hostModules) return hostModules;

  if (currentHost === "powerpoint") {
    const [pptContext, pptSnapshot, pptOps] = await Promise.all([
      import("./ppt/context.js"),
      import("./ppt/snapshot.js"),
      import("./ppt/operations.js"),
    ]);
    hostModules = {
      context: pptContext,
      snapshot: pptSnapshot,
      ops: pptOps,
    };
  } else {
    const [excelContext, excelSnapshot] = await Promise.all([
      import("./excel/context.js"),
      import("./excel/snapshot.js"),
    ]);
    hostModules = {
      context: excelContext,
      snapshot: excelSnapshot,
      ops: null,
    };
  }

  return hostModules;
}

function getHostContextFn() {
  return currentHost === "powerpoint"
    ? hostModules.context.getSlideContext
    : hostModules.context.getSheetContext;
}

function getHostFullContextFn() {
  return currentHost === "powerpoint"
    ? hostModules.context.readFullPresentationContext
    : hostModules.context.readFullWorkbookContext;
}

function getHostMetadataFn() {
  return currentHost === "powerpoint"
    ? hostModules.context.readPresentationMetadata
    : hostModules.context.readWorkbookMetadata;
}

function getHostCaptureSnapshotFn() {
  return currentHost === "powerpoint"
    ? hostModules.snapshot.captureSnapshot
    : hostModules.snapshot.captureSnapshot;
}

function getHostVerifyFn() {
  return currentHost === "powerpoint"
    ? hostModules.snapshot.verifyPptOperations
    : hostModules.snapshot.verifyOperations;
}

function getHostExtractNamesFn() {
  return currentHost === "powerpoint"
    ? hostModules.snapshot.extractSlideNamesFromResults
    : hostModules.snapshot.extractSheetNamesFromResults;
}

// ─── State ───────────────────────────────────────────────────────────────────

let conversationHistory = [];
let isExecuting = false;
let isStopped = false;
let currentMode = localStorage.getItem("agentMode") || "interactive";
let pendingImageBase64 = null;

// Step tracking for retry/revert buttons
let stepStack = []; // [{stepNumber, plan, operations, resultEl, hadError, feedback, snapshotBefore}]

// Stores the user's original query for validation before declaring completion
let originalQuery = "";

// ─── Feedback State ─────────────────────────────────────────────────────────

let feedbackStepData = null;
let feedbackTotalSteps = 0;
let feedbackTotalTiming = 0;

// ─── Stop Handler ────────────────────────────────────────────────────────────

export function stopAgent() {
  isStopped = true;
  isExecuting = false;
  console.log("[agentChat] Agent stopped by user");
}

// ─── Mode Switching ──────────────────────────────────────────────────────────

export function switchMode(mode) {
  currentMode = switchModeUI(mode, appendMessage) || currentMode;
}

// ─── Clear / Reset ───────────────────────────────────────────────────────────

export function clearChat() {
  // Kill any running AI request immediately
  isStopped = true;
  abortActiveRequest();

  conversationHistory = [];
  stepStack = [];
  isExecuting = false;
  isStopped = false;
  feedbackStepData = null;
  feedbackTotalSteps = 0;
  feedbackTotalTiming = 0;
  originalQuery = "";

  flushTelemetry(false);
  clearChatUI(showWelcome);
  resetStepGroup();

  console.log("[agentChat] Chat cleared");
}

function resetForNextTask() {
  conversationHistory = [];
  stepStack = [];
  originalQuery = "";
  feedbackStepData = null;
  feedbackTotalSteps = 0;
  feedbackTotalTiming = 0;

  flushTelemetry(true);
  initTelemetry();
  clearChatUI(showWelcome);
  resetStepGroup();
}

// ─── Retry / Revert ─────────────────────────────────────────────────────────

async function handleRetry(stepIndex) {
  if (stepIndex < 0 || stepIndex >= stepStack.length) return;
  const step = stepStack[stepIndex];

  const opSummary = step.operations
    .map((o) => `${o.operation}(${o.sheet}, ${o.range || o.name || ""})`)
    .join(", ");

  let retryContent = `I'm not satisfied with the previous action in Step ${step.stepNumber}. `;
  if (step.hadError) {
    retryContent += `The operations failed with errors:\n${step.feedback}\n\n`;
  } else {
    retryContent += `The result was not what I wanted.\n`;
  }
  retryContent += `Operations that were attempted: ${opSummary}\n`;
  retryContent += `Plan: ${step.plan}\n`;
  retryContent += `Please retry with corrections — fix the issues and provide updated structured operations.`;

  conversationHistory.push({ role: "user", content: retryContent });
  stepStack.splice(stepIndex, stepStack.length - stepIndex);

  if (step.resultEl && step.resultEl.parentElement) {
    step.resultEl.parentElement.removeChild(step.resultEl);
  }

  await continueExecutionFromAI();
}

async function handleRevert(stepIndex) {
  if (stepIndex < 0 || stepIndex >= stepStack.length) return;
  const step = stepStack[stepIndex];

  const opSummary = step.operations
    .map((o) => `${o.operation}(${o.sheet}, ${o.range || o.name || ""})`)
    .join(", ");

  const revertContent =
    `I want to undo/revert the action from Step ${step.stepNumber}. ` +
    `Operations that were executed: ${opSummary}\n` +
    `Plan: ${step.plan}\n` +
    `Please undo these changes. If you have access to the previous state, restore it. ` +
    `Otherwise, reverse each operation manually.`;

  conversationHistory.push({ role: "user", content: revertContent });
  stepStack.splice(stepIndex, stepStack.length - stepIndex);

  if (step.resultEl && step.resultEl.parentElement) {
    step.resultEl.parentElement.removeChild(step.resultEl);
  }

  await continueExecutionFromAI();
}

// ─── Improve Flow ────────────────────────────────────────────────────────────

async function handleImprove(comment, taskStats) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer || !feedbackStepData) return;

  const improveArea = messagesContainer.querySelector(".agent-improve-area");
  if (improveArea) improveArea.style.display = "none";

  const buttons = messagesContainer.querySelectorAll(".agent-feedback-btn");
  buttons.forEach((btn) => (btn.disabled = true));

  appendMessage("agent", `🔧 Applying improvements: "${comment}"`);

  await loadHostModules();
  const sheetContext = await getHostContextFn()();
  const systemPrompt = buildSystemPrompt(currentHost, currentMode, sheetContext);

  const opField = currentHost === "powerpoint" ? "slide" : "sheet";
  const feedbackMessages = [
    { role: "system", content: systemPrompt },
    ...conversationHistory,
    {
      role: "user",
      content: `The user wants to improve the previous result.

Original request: "${originalQuery}"

Previous operations executed:
${feedbackStepData.operations.map((o) => `- ${o.operation}${o[opField] ? ` on ${o[opField]}` : ""}${o.range ? ` (${o.range})` : ""}`).join("\n")}

User feedback: "${comment}"

Please generate new structured operations to incorporate this improvement. Output inside a JSON code block.`,
    },
  ];

  const improveResponse = await streamFromAI(feedbackMessages, () => isStopped);
  if (!improveResponse || isStopped) {
    buttons.forEach((btn) => (btn.disabled = false));
    appendMessage("agent", "⚠ Improvement cancelled.");
    return;
  }

  conversationHistory.push({ role: "assistant", content: improveResponse });
  let improveOps = extractOperations(improveResponse);

  if (!improveOps || improveOps.operations.length === 0) {
    if (isActionableResponse(improveResponse)) {
      appendMessage("agent", "⚠ Could not parse operations. Asking again...");
      const retryMessages = [
        { role: "system", content: systemPrompt },
        ...conversationHistory,
        {
          role: "user",
          content: `Please provide your improvement as structured operations inside a JSON code block.`,
        },
      ];
      const retryResponse = await streamFromAI(retryMessages, () => isStopped);
      if (!retryResponse) {
        buttons.forEach((btn) => (btn.disabled = false));
        return;
      }
      conversationHistory.push({ role: "assistant", content: retryResponse });
      const retryOps = extractOperations(retryResponse);
      if (!retryOps || retryOps.operations.length === 0) {
        appendMessage("agent", `ℹ ${improveResponse.substring(0, 300)}`);
        submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
        resetForNextTask();
        return;
      }
      improveOps = retryOps;
    } else {
      appendMessage("agent", `ℹ ${improveResponse.substring(0, 300)}`);
      submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
      resetForNextTask();
      return;
    }
  }

  console.log("[improve] Executing", improveOps.operations.length, "improvement operations");
  appendMessage(
    "agent",
    `🔄 Executing ${improveOps.operations.length} improvement operation(s)...`
  );

  const improveResults = [];
  let improveHadError = false;
  const improveStepStart = Date.now();

  for (const op of improveOps.operations) {
    try {
      const runFn = currentHost === "powerpoint" ? PowerPoint.run : Excel.run;
      const result = await runFn(async (context) => {
        return executeOperationWithTracking(context, op);
      });
      improveResults.push(result);
    } catch (error) {
      improveResults.push(`✗ ${op.operation} failed: ${error.message}`);
      improveHadError = true;

      // Record error in telemetry
      try {
        const { recordOpResult } = await import("./telemetry.js");
        recordOpResult(op.operation, false, error);
      } catch {
        // Non-critical
      }
    }
  }

  const improveTiming = Date.now() - improveStepStart;

  if (improveHadError) {
    appendMessage("agent", `❌ Some improvements failed: ${improveResults.join("\n")}`);
    submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
    resetForNextTask();
    return;
  }

  const actualNames = getHostExtractNamesFn()(improveResults, improveOps.operations);
  const improvementVerification = await getHostVerifyFn()(improveOps.operations, actualNames);
  appendMessage(
    "agent",
    `✅ Improvements applied.${improvementVerification ? "\n" + improvementVerification : ""}`
  );

  trackStep({
    stepNumber: feedbackTotalSteps + 1,
    userPrompt: originalQuery,
    plan: `Improvement: ${comment}`,
    operations: improveOps.operations,
    results: improveResults,
    errors: [],
    success: true,
    timingMs: improveTiming,
    aiResponse: improveResponse,
    verification: improvementVerification || "",
    conversationHistory: [...conversationHistory],
    sheetContext,
    systemPrompt,
  });
  feedbackTotalSteps++;
  feedbackTotalTiming += improveTiming;

  try {
    const afterContext = await getHostMetadataFn()();
    if (afterContext) {
      trackSnapshot("after", afterContext);
    }
  } catch {
    // Non-critical
  }

  // Show feedback buttons
  const lastBubble = messagesContainer.querySelector(
    ".agent-message.agent:last-child .agent-bubble"
  );
  if (lastBubble) {
    appendFeedbackButtons(lastBubble, {
      onOk: () => {
        submitFeedback("ok", comment, { ...taskStats, improvementApplied: true });
        resetForNextTask();
      },
      onBad: () => {
        submitFeedback("bad", comment, { ...taskStats, improvementApplied: true });
        resetForNextTask();
      },
      onImprove: (newComment) => {
        handleImprove(newComment, { ...taskStats, improvementApplied: true });
      },
    });
  }

  scrollChatToBottom();
}

// ─── Continue Execution (after retry/revert) ────────────────────────────────

async function continueExecutionFromAI() {
  await loadHostModules();
  const sheetContext = await getHostContextFn()();
  const systemPrompt = buildSystemPrompt(currentHost, currentMode, sheetContext);

  const nextApiMessages = [
    { role: "system", content: systemPrompt },
    ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
  ];

  const nextResponse = await streamFromAI(nextApiMessages, () => isStopped);
  if (!nextResponse || isStopped) return;

  conversationHistory.push({ role: "assistant", content: nextResponse });
  const nextOps = extractOperations(nextResponse);
  if (!nextOps || nextOps.operations.length === 0) {
    if (isCompletionClaim(nextResponse)) {
      console.log("[validate] AI claims complete (retry path) — validating...");
      appendMessage("agent", "🔍 Verifying completion...");

      const fullContext = await getHostFullContextFn()();
      const hostLabel = currentHost === "powerpoint" ? "presentation" : "workbook";
      const validationMessages = [
        {
          role: "system",
          content: `You are a ${currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task validator. Your ONLY job is to check whether the agent completed the user's request. You are NOT an agent — you do NOT generate operations.

USER REQUEST: "${originalQuery}"

CURRENT ${hostLabel.toUpperCase()} STATE:

${fullContext}

## VALIDATION RULES

1. Break the user's request into specific requirements. Check each one against the current state.
2. ONLY flag issues you can point to in the data. Do NOT say "I cannot see visual formatting" — if conditional formatting was requested, check that the correct formulas/ranges exist. If a chart was requested, check that a chart object exists.
3. Be strict but fair. Missing sheets, wrong formulas, empty data, incorrect aggregations = real issues. Minor formatting = not an issue.
4. If validated=true, issues MUST be empty. If validated=false, issues MUST list every problem.

## STRICT OUTPUT RULE

You MUST respond with ONLY a JSON object. NO explanations, NO code blocks, NO operations.

Respond with JSON:
{
  "validated": true/false,
  "issues": ["list of specific issues or empty array"],
  "summary": "Brief summary"
}

FORBIDDEN: Do NOT include any "operations" array. Do NOT generate any operations. Do NOT suggest new steps. Your role is validation ONLY.

CRITICAL: Do NOT pass validation unless each requirement is verified.`,
        },
        ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
      ];

      const validationResult = await streamFromAI(validationMessages, () => isStopped);
      if (validationResult) {
        const jsonMatch = validationResult.match(/```(?:json)?\s*[\r\n]+([\s\S]*?)```/);
        let parsedResult;
        try {
          parsedResult = jsonMatch
            ? JSON.parse(jsonMatch[1].trim())
            : JSON.parse(validationResult.trim());
        } catch {
          parsedResult = { validated: true, issues: [], summary: validationResult };
        }

        if (parsedResult.validated && parsedResult.issues.length === 0) {
          appendMessage("agent", `✅ ${parsedResult.summary || "Verified and complete!"}`);
          conversationHistory = [];
          stepStack = [];
          originalQuery = "";
          return;
        } else {
          const issuesText =
            parsedResult.issues.length > 0
              ? parsedResult.issues.join(", ")
              : "Validation inconclusive";
          appendMessage("agent", `⚠ Issues: ${issuesText}. Asking agent to fix...`);

          const fixMessages = [
            { role: "system", content: systemPrompt },
            ...conversationHistory,
            {
              role: "user",
              content: `⚠ Validation failed: ${parsedResult.issues.join("\n")}. Please fix with structured operations.`,
            },
          ];

          const fixResponse = await streamFromAI(fixMessages, () => isStopped);
          if (!fixResponse || isStopped) return;

          conversationHistory.push({ role: "assistant", content: fixResponse });
          const fixOps = extractOperations(fixResponse);
          if (fixOps && fixOps.operations.length > 0) {
            await executeOperationsLoop(fixOps, sheetContext, systemPrompt);
            return;
          } else {
            appendMessage("agent", fixResponse.substring(0, 300));
            return;
          }
        }
      } else {
        appendMessage("agent", "⚠ Validation skipped.");
        conversationHistory = [];
        stepStack = [];
        originalQuery = "";
        return;
      }
    }
    appendMessage("agent", "⚠ No operations found. Asking to continue...");
    const retryFeedback =
      "Your last response did not contain structured operations. Please provide the next step as structured operations, or say 'All steps complete.' if done.";
    conversationHistory.push({ role: "user", content: retryFeedback });
    const retryResponse = await streamFromAI(
      [
        { role: "system", content: systemPrompt },
        ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
      ],
      () => isStopped
    );
    if (!retryResponse || isStopped) return;
    conversationHistory.push({ role: "assistant", content: retryResponse });
    const retryOps = extractOperations(retryResponse);
    if (!retryOps || retryOps.operations.length === 0) {
      appendMessage("agent", "All steps complete.");
      conversationHistory = [];
      stepStack = [];
      originalQuery = "";
      return;
    }
    await executeOperationsLoop(retryOps, sheetContext, systemPrompt);
    return;
  }
  await executeOperationsLoop(nextOps, sheetContext, systemPrompt);
}

// ─── Operations Execution Loop ───────────────────────────────────────────────

async function executeOperationsLoop(initialOps, initialContext, systemPrompt) {
  const plan = initialOps?.plan || "";
  const opsArray = Array.isArray(initialOps) ? initialOps : initialOps?.operations || [];
  let validOps = [...opsArray];
  let sheetContext = initialContext;
  let stepNumber = 1;

  while (validOps.length > 0 && !isStopped) {
    const stepStartTime = Date.now();
    appendStepMessage(stepNumber, plan, validOps);
    stepNumber++;

    const snapshotBefore = await getHostCaptureSnapshotFn()(validOps);

    const results = [];
    let hadError = false;

    for (const op of validOps) {
      try {
        const runFn = currentHost === "powerpoint" ? PowerPoint.run : Excel.run;
        const result = await runFn(async (context) => {
          return executeOperationWithTracking(context, op);
        });
        results.push(result);
      } catch (error) {
        results.push(`✗ ${op.operation} failed: ${error.message}`);
        hadError = true;

        // Record error in telemetry
        try {
          const { recordOpResult } = await import("./telemetry.js");
          recordOpResult(op.operation, false, error);
        } catch {
          // Non-critical
        }
      }
    }

    const resultText = results.join("\n");
    const resultEl = appendExecutionResult({
      success: !hadError,
      result: resultText,
      stepNumber: stepNumber - 1,
      operations: validOps,
      hadError,
      _onRetry: handleRetry,
      _onRevert: handleRevert,
    });

    // Feed execution result back to AI
    let feedback;
    let verification = "";
    if (hadError) {
      feedback = `Some operations failed: ${resultText}. Adapt and provide corrected operations.`;
      console.log(`[step:${stepNumber}] Feedback sent (error):`, feedback);
      conversationHistory.push({ role: "user", content: feedback });
    } else {
      const opField = currentHost === "powerpoint" ? "slide" : "sheet";
      const actualNames = getHostExtractNamesFn()(results, validOps);
      verification = await getHostVerifyFn()(validOps, actualNames);
      const opSummary = validOps
        .map((o) => `${o.operation}(${o[opField] || o.name || ""}, ${o.range || o.shape || ""})`)
        .join(", ");
      const resultsSection = results.length > 0 ? `\n\nOperation results:\n${resultText}` : "";
      const verificationSection = verification ? `\n\nVerification:\n${verification}` : "";
      feedback = `Step executed. Operations: ${opSummary}${resultsSection}${verificationSection}\n\nCurrent ${currentHost === "powerpoint" ? "slide" : "sheet"} state is in your system prompt. Provide the next step as structured operations, or say "All steps complete." if done.`;
      console.log(`[step:${stepNumber}] Feedback sent (success):`, feedback);
      conversationHistory.push({ role: "user", content: feedback });
    }

    stepStack.push({
      stepNumber: stepNumber - 1,
      operations: validOps,
      resultEl,
      hadError,
      feedback,
      snapshotBefore,
      plan,
    });

    trackStep({
      stepNumber: stepNumber - 1,
      userPrompt: originalQuery,
      plan,
      operations: validOps,
      results: results,
      errors: hadError ? [resultText] : [],
      success: !hadError,
      timingMs: Date.now() - stepStartTime,
      aiResponse: nextResponse || "",
      verification: verification || "",
      conversationHistory: [...conversationHistory],
      sheetContext,
      systemPrompt,
    });

    const stepTiming = Date.now() - stepStartTime;
    feedbackTotalSteps++;
    feedbackTotalTiming += stepTiming;

    // Read updated context
    sheetContext = await getHostContextFn()();
    console.log(`[step:${stepNumber}] Updated context:`, sheetContext);
    systemPrompt = buildSystemPrompt(currentHost, currentMode, sheetContext);

    const nextApiMessages = [
      { role: "system", content: systemPrompt },
      ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
    ];
    console.log(
      `[step:${stepNumber}] Sending to AI — msg count:`,
      nextApiMessages.length,
      "system prompt length:",
      systemPrompt.length
    );

    const nextResponse = await streamFromAI(nextApiMessages, () => isStopped);
    console.log(
      `[step:${stepNumber}] AI response received (length:`,
      nextResponse ? nextResponse.length : 0,
      `):\n`,
      nextResponse
    );

    if (!nextResponse || isStopped) {
      if (isStopped) {
        appendMessage("agent", "⏹ Execution stopped.");
      } else {
        appendMessage("agent", "All steps complete.");
      }
      if (!isStopped) {
        conversationHistory = [];
        stepStack = [];
        originalQuery = "";
      }
      break;
    }

    conversationHistory.push({ role: "assistant", content: nextResponse });
    const nextOps = extractOperations(nextResponse);
    console.log(`[step:${stepNumber}] Extracted ops:`, nextOps);
    if (!nextOps || nextOps.operations.length === 0) {
      if (isCompletionClaim(nextResponse)) {
        console.log(
          "[validate] AI claims complete — validating against full",
          currentHost,
          "state..."
        );
        appendMessage("agent", `🔍 Verifying completion against full ${currentHost} state...`);

        const fullContext = await getHostFullContextFn()();
        console.log("[validate] Full context:\n", fullContext);

        const hostLabel = currentHost === "powerpoint" ? "presentation" : "workbook";
        const opField = currentHost === "powerpoint" ? "slide" : "sheet";
        const opsList = validOps
          .map(
            (o) =>
              `- ${o.operation}${o[opField] ? ` on ${o[opField]}` : ""}${o.range ? ` (${o.range})` : ""}`
          )
          .join("\n");
        const validationMessages = [
          {
            role: "system",
            content: `You are a ${currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task validator. Your ONLY job is to check whether the agent completed the user's request. You are NOT an agent — you do NOT generate operations.

USER REQUEST: "${originalQuery}"

OPERATIONS EXECUTED:
${opsList}

CURRENT ${hostLabel.toUpperCase()} STATE:

${fullContext}

## VALIDATION RULES

1. Break the user's request into specific requirements. Check each one against the current state.
2. ONLY flag issues you can point to in the data above. Do NOT say "I cannot see visual formatting" — if conditional formatting was requested, check that the correct formulas/ranges exist. If a chart was requested, check that a chart object exists.
3. Be strict but fair. Missing sheets, wrong formulas, empty data, incorrect aggregations = real issues. Minor formatting = not an issue.
4. If validated=true, issues MUST be empty. If validated=false, issues MUST list every problem.

## STRICT OUTPUT RULE

You MUST respond with ONLY a JSON object. NO explanations, NO code blocks, NO operations.

Respond with JSON:
{
  "validated": true/false,
  "issues": ["list of specific issues or empty array"],
  "summary": "Brief summary"
}

FORBIDDEN: Do NOT include any "operations" array. Do NOT generate any operations. Do NOT suggest new steps. Your role is validation ONLY.

CRITICAL: Do NOT pass validation unless each requirement is verified. The user will be frustrated if issues are missed.`,
          },
          ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
        ];

        const validationResult = await streamFromAI(validationMessages, () => isStopped);
        console.log("[validate] AI validation response:\n", validationResult);

        if (validationResult) {
          const jsonMatch = validationResult.match(/```(?:json)?\s*[\r\n]+([\s\S]*?)```/);
          let parsedResult;
          try {
            if (jsonMatch) {
              parsedResult = JSON.parse(jsonMatch[1].trim());
            } else {
              parsedResult = JSON.parse(validationResult.trim());
            }
          } catch {
            parsedResult = { validated: true, issues: [], summary: validationResult };
          }

          if (parsedResult.validated && parsedResult.issues.length === 0) {
            console.log("[validate] Validation passed — generating summary...");

            try {
              const afterContext = await getHostMetadataFn()();
              if (afterContext) {
                trackSnapshot("after", afterContext);
              }
            } catch {
              // Non-critical
            }

            const opField = currentHost === "powerpoint" ? "slide" : "sheet";
            feedbackStepData = {
              userPrompt: originalQuery,
              plan,
              operations: validOps,
              results: results,
              errors: [],
              success: true,
              timingMs: feedbackTotalTiming,
            };

            trackStep({
              stepNumber: feedbackTotalSteps,
              userPrompt: originalQuery,
              plan,
              operations: validOps,
              results: results,
              errors: [],
              success: true,
              timingMs: feedbackTotalTiming,
              aiResponse: nextResponse || "",
              verification: verification || "",
              conversationHistory: [...conversationHistory],
              sheetContext,
              systemPrompt,
            });

            const taskStats = {
              totalSteps: feedbackTotalSteps,
              totalTiming: feedbackTotalTiming,
              finalSuccess: true,
            };

            // Generate summary
            const summaryMessages = [
              {
                role: "system",
                content: `You are a ${currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task summarizer. Based on the operations performed and the final ${currentHost === "powerpoint" ? "presentation" : "workbook"} state, provide a clear, user-friendly summary of what was accomplished.

Respond with a JSON object:
{
  "title": "Brief title of what was done",
  "changes": ["list of specific changes made, one per line"],
  "details": "More detailed explanation of the changes and their impact"
}`,
              },
              {
                role: "user",
                content: `Original request: "${originalQuery}"

Operations performed:
${validOps.map((op) => `- ${op.operation} on ${op[opField] || op.name}${op.range ? ` (${op.range})` : ""}`).join("\n")}

Final ${currentHost === "powerpoint" ? "presentation" : "workbook"} state:
${fullContext}

Provide a clear summary of what was done.`,
              },
            ];

            let summaryText = "";
            try {
              const summaryResponse = await streamFromAI(summaryMessages, () => isStopped);
              if (summaryResponse) {
                const jsonMatch2 = summaryResponse.match(/```(?:json)?\s*[\r\n]+([\s\S]*?)```/);
                try {
                  if (jsonMatch2) {
                    summaryText = JSON.parse(jsonMatch2[1].trim());
                  } else {
                    summaryText = JSON.parse(summaryResponse.trim());
                  }
                } catch {
                  summaryText = {
                    title: parsedResult.summary || "Task complete",
                    changes: [],
                    details: summaryResponse,
                  };
                }
              }
            } catch (e) {
              console.warn("[summary] Failed to generate summary:", e.message);
              summaryText = {
                title: parsedResult.summary || "Task complete",
                changes: [],
                details: "",
              };
            }

            // Show completion with feedback
            const messagesContainer = document.getElementById("chat-messages");
            if (messagesContainer) {
              const summaryEl = buildCompletionSummary(summaryText, escapeHtml);
              messagesContainer.appendChild(summaryEl);
              scrollChatToBottom();

              const bubbleEl = summaryEl.querySelector(".agent-bubble");
              appendFeedbackButtons(bubbleEl, {
                onOk: () => {
                  submitFeedback("ok", "", taskStats);
                  console.log("[telemetry] User rated: OK");
                  resetForNextTask();
                },
                onBad: () => {
                  submitFeedback("bad", "", taskStats);
                  console.log("[telemetry] User rated: Bad");
                  resetForNextTask();
                },
                onImprove: (comment) => {
                  handleImprove(comment, taskStats);
                },
              });
            }

            // Don't reset yet — wait for feedback
            break;
          } else {
            const issuesText =
              parsedResult.issues.length > 0
                ? `Issues found: ${parsedResult.issues.join(", ")}`
                : "Validation inconclusive";
            appendMessage("agent", `⚠ ${issuesText}. Asking agent to fix...`);
            console.log("[validate] Validation failed — sending issues back to agent");

            const fixMessages = [
              { role: "system", content: systemPrompt },
              ...conversationHistory,
              {
                role: "user",
                content: `⚠ Validation failed. The following issues were found:\n${parsedResult.issues.join("\n")}\n\nPlease fix these issues with structured operations, or if everything is actually correct, explain why in plain text.`,
              },
            ];

            const fixResponse = await streamFromAI(fixMessages, () => isStopped);
            console.log("[validate] Fix response received:", fixResponse ? `${fixResponse.length} chars` : "null/empty");
            if (!fixResponse || isStopped) {
              if (isStopped) {
                appendMessage("agent", "⏹ Execution stopped.");
              } else {
                appendMessage("agent", "⚠ Agent did not respond to the fix request. Try again manually.");
              }
              break;
            }

            conversationHistory.push({ role: "assistant", content: fixResponse });
            const fixOps = extractOperations(fixResponse);
            console.log("[validate] Fix ops extracted:", fixOps ? fixOps.operations.length : 0, "operations");
            if (fixOps && fixOps.operations.length > 0) {
              validOps.length = 0;
              validOps.push(...fixOps.operations);
              console.log(
                "[validate] Agent received fix request with",
                fixOps.operations.length,
                "operations"
              );
              // Continue the while loop to execute fix operations
              continue;
            } else {
              appendMessage("agent", `ℹ ${fixResponse.substring(0, 300)}`);
              trackStep({
                stepNumber: stepNumber - 1,
                userPrompt: originalQuery,
                plan,
                operations: validOps,
                results: results,
                errors: ["Validation fix AI responded with plain text instead of operations"],
                success: false,
                timingMs: 0,
                aiResponse: fixResponse || "(null)",
                verification: verification || "",
                conversationHistory: [...conversationHistory],
                sheetContext,
                systemPrompt,
              });
              break;
            }
          }
        } else {
          appendMessage("agent", "⚠ Validation skipped — no response from validator.");
          conversationHistory = [];
          stepStack = [];
          originalQuery = "";
          break;
        }
        break;
      }

      // No operations — ask for clarification
      appendMessage("agent", "⚠ No operations found. Asking to continue...");
      const retryFeedback =
        "Your last response did not contain structured operations. Please provide the next step as structured operations, or say 'All steps complete.' if done.";
      conversationHistory.push({ role: "user", content: retryFeedback });

      const retryResponse = await streamFromAI(
        [
          { role: "system", content: systemPrompt },
          ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
        ],
        () => isStopped
      );

      if (!retryResponse) break;

      conversationHistory.push({ role: "assistant", content: retryResponse });
      const retryOps = extractOperations(retryResponse);
      if (!retryOps || retryOps.operations.length === 0) {
        appendMessage("agent", "⚠ Agent could not generate operations. Stopping.");
        break;
      }

      validOps.length = 0;
      validOps.push(...retryOps.operations);
    } else {
      validOps.length = 0;
      validOps.push(...nextOps.operations);
    }
  }

  finalizeStepGroup();
}

// ─── Initialization ──────────────────────────────────────────────────────────

export async function initAgentChat() {
  try {
    console.log("[agentChat] initAgentChat called");

    // Ensure host is set
    if (!localStorage.getItem("agentHost")) {
      localStorage.setItem("agentHost", currentHost);
    }
    currentHost = localStorage.getItem("agentHost") || "excel";

    const appBody = document.getElementById("app-body");
    console.log("[agentChat] appBody:", appBody, "host:", currentHost);
    if (!appBody) {
      console.error("[agentChat] #app-body not found in DOM");
      return;
    }

    // Fetch manifest version and store in localStorage
    try {
      const manifestRes = await fetch("manifest.xml");
      const manifestText = await manifestRes.text();
      const versionMatch = manifestText.match(/<Version>([^<]+)<\/Version>/);
      if (versionMatch) {
        const manifestVersion = versionMatch[1].replace(/\.\d+$/, "");
        localStorage.setItem("manifestVersion", manifestVersion);
        console.log("[agentChat] Manifest version:", manifestVersion);
      }
    } catch {
      console.log("[agentChat] Could not fetch manifest.xml");
    }

    // Create chat container
    const chatContainer = document.createElement("div");
    chatContainer.className = "agent-chat-container";
    chatContainer.id = "agent-chat";
    chatContainer.style.display = "none";

    const aiConfig = getConfig();
    const version = aiConfig.version || "1.0.0";

    // Check manifest version vs app version
    const storedManifestVersion = localStorage.getItem("manifestVersion");
    if (storedManifestVersion && storedManifestVersion !== version) {
      const versionBanner = document.getElementById("version-banner");
      const versionBannerText = document.getElementById("version-banner-text");
      if (versionBanner && versionBannerText) {
        versionBannerText.textContent = `Manifest version (${storedManifestVersion}) differs from current VERSION (${version})`;
        versionBanner.style.display = "block";
      }
    }

    chatContainer.innerHTML = `
        <div class="agent-mode-bar">
          <span id="mode-badge" style="font-size:12px;font-weight:600;">${currentMode === "interactive" ? "⚡ Edit Mode" : "📖 Read Only"}</span>
          <div class="agent-mode-toggle">
            <button id="mode-interactive" class="agent-mode-btn ${currentMode === "interactive" ? "mode-active" : ""}">⚡ Edit Mode</button>
            <button id="mode-explain" class="agent-mode-btn ${currentMode === "explain" ? "mode-active" : ""}">📖 Read Only</button>
          </div>
        </div>
        <div class="agent-version-banner" id="version-banner" style="display:none; padding:4px 12px; background:#f0f6ff; border-bottom:1px solid #d0e3f5; font-size:11px; color:#5f6b7a; text-align:center;">
          <span id="version-banner-text"></span>
        </div>
        <div class="agent-config" id="agent-config">
          <button id="config-toggle" class="config-toggle" title="Toggle settings">
            <span class="toggle-icon">⚙️</span>
            <span class="toggle-text" id="toggle-text">Settings</span>
          </button>
          <div class="config-fields">
            <div class="config-field">
              <label for="agent-endpoint">Endpoint:</label>
              <input type="text" id="agent-endpoint" value="${escapeHtml(aiConfig.endpoint)}" placeholder="http://localhost:8081/v1/chat/completions" />
            </div>
            <div class="config-field">
              <label for="agent-apikey">API Key:</label>
              <input type="password" id="agent-apikey" value="${escapeHtml(aiConfig.apiKey)}" placeholder="Bearer token (optional)" />
            </div>
            <div class="config-field">
              <label for="agent-model">Model:</label>
              <input type="text" id="agent-model" list="model-suggestions" value="${escapeHtml(aiConfig.model)}" placeholder="e.g. nvidia/nemotron-3-super-120b-a12b:free" />
              <div id="model-status"></div>
              <datalist id="model-suggestions">
                <option value="nvidia/nemotron-3-super-120b-a12b:free">
                <option value="openai/gpt-4">
                <option value="openai/gpt-4o">
                <option value="google/gemini-2.5-pro-preview-05-06">
                <option value="meta-llama/llama-3.3-70b-instruct">
                <option value="mistralai/mistral-large-2-instruct">
                <option value="deepseek/deepseek-chat">
                <option value="nousresearch/hermes-3-llama-3.1-70b">
                <option value="local-model">
              </datalist>
            </div>
            <div class="config-field" style="flex-direction:row; align-items:center; justify-content:space-between; padding:4px 0;">
              <label style="font-size:11px; font-weight:600; color:var(--chat-text-secondary); text-transform:uppercase; letter-spacing:0.03em; margin:0;">Telemetry</label>
              <label style="display:flex; align-items:center; gap:8px; cursor:pointer;">
                <span id="telemetry-status" style="font-size:11px; color:var(--chat-text-muted);">On</span>
                <input type="checkbox" id="telemetry-toggle" ${localStorage.getItem("telemetryEnabled") !== "false" ? "checked" : ""} style="width:16px; height:16px; cursor:pointer;" />
              </label>
            </div>
            <div class="config-field">
              <label for="agent-version">Version:</label>
              <input type="text" id="agent-version" value="${escapeHtml(version)}" placeholder="1.0.0" readonly style="opacity: 0.7; cursor: not-allowed;" />
            </div>
          </div>
        </div>
      <div class="agent-messages" id="chat-messages">
        ${showWelcome ? "" : ""}
        <div class="agent-welcome">
          <h3><img src="assets/icon-80.svg" alt="KuroAgent" style="width:28px;height:28px;vertical-align:middle;margin-right:8px;border-radius:6px;">Kuro — Your Excel Copilot</h3>
          <p style="font-size: 13px; color: #666; margin-bottom: 16px;">
            Talk to your spreadsheet in plain English. Kuro reads your data, plans the steps, and does the work for you.
          </p>
          <p>${currentHost === "powerpoint" ? "Describe what you want to do in PowerPoint. The agent will generate and execute structured operations automatically." : "Tell Kuro what you need — it figures out the rest."}</p>
          <p style="margin-top: 12px; font-size: 12px; color: #999;">Examples:<br/>
            ${
              currentHost === "powerpoint"
                ? '"Add a slide titled Q4 Results"<br/>' +
                  '"Set text on Title 1 to Welcome"<br/>' +
                  '"Add a rectangle shape on Slide 2"'
                : '"Put Hello in A1, make it bold"<br/>' +
                  '"Chart the data in A1:B5"<br/>' +
                  '"Sum up column A, results in C"'
            }</p>
          <div class="agent-footer">
            <p style="font-size: 11px; color: #999; margin: 16px 0 8px 0; font-weight: 600;">About</p>
            <p style="font-size: 12px; color: #666; margin: 0 0 8px 0;">
              Created by <strong>Yann Loic Assoumou</strong>
            </p>
            <p style="font-size: 12px; color: #666; margin: 0;">
              <a href="https://github.com/yannassoumou" target="_blank" rel="noopener" style="color: #0078d4; text-decoration: none; margin-right: 12px;">📦 GitHub</a>
              <a href="https://www.linkedin.com/in/anghaiassoumou/" target="_blank" rel="noopener" style="color: #0078d4; text-decoration: none;">💼 LinkedIn</a>
            </p>
          </div>
        </div>
      </div>
      <div class="agent-input-area">
        <textarea id="chat-input" placeholder="Ask Kuro anything about your ${currentHost === "powerpoint" ? "slides" : "spreadsheet"}..." rows="1"></textarea>
        <input type="file" id="chat-image-input" accept="image/*" style="display:none;" />
        <button id="chat-image-upload" title="Attach image" style="display:none;">📎</button>
        <button id="chat-send" disabled>Send</button>
        <button id="chat-stop" style="display:none; background:#dc3545; color:white; border:none; padding:8px 16px; border-radius:4px; cursor:pointer; font-size:14px;">Stop</button>
        <button id="chat-clear" class="clear-btn" title="Clear chat">🗑</button>
      </div>
    `;

    document.body.appendChild(chatContainer);

    // Restyle app-body with new welcome content
    const host = currentHost === "powerpoint" ? "PowerPoint" : "Excel";
    const hostLower = currentHost === "powerpoint" ? "PowerPoint" : "Excel";
    appBody.innerHTML = `
      <div class="welcome-card-inner">
        <div class="welcome-card-icon"><img src="assets/icon-80.svg" alt="KuroAgent" style="width:48px;height:48px;border-radius:10px;"></div>
        <h2>Kuro — Your ${host} Copilot</h2>
        <p class="welcome-subtitle">Talk to your ${hostLower === "PowerPoint" ? "presentations" : "spreadsheets"} in plain English. Kuro reads your data, plans the steps, and does the work for you.</p>
        <p class="welcome-description">Tell Kuro what you need — it figures out the rest.</p>
        <p class="welcome-section-label">Examples</p>
        <div class="welcome-examples">
          <ul>${
            hostLower === "PowerPoint"
              ? '<li>"Add a slide titled Q4 Results"</li><li>"Set text on Title 1 to Welcome"</li><li>"Add a rectangle shape on Slide 2"</li>'
              : '<li>"Put Hello in A1, make it bold"</li><li>"Chart the data in A1:B5"</li><li>"Sum up column A, results in C"</li>'
          }</ul>
        </div>
        <div class="welcome-footer">
          <p class="welcome-creator">Created by <strong>Yann Loic Assoumou</strong></p>
          <div class="welcome-links">
            <a href="https://github.com/yannassoumou" target="_blank" rel="noopener">📦 GitHub</a>
            <a href="https://www.linkedin.com/in/anghaiassoumou/" target="_blank" rel="noopener">💼 LinkedIn</a>
          </div>
        </div>
        <button id="open-chat-btn" style="margin-top: 24px; padding: 12px 32px; background: var(--chat-accent, #0078d4); color: #fff; border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; font-family: inherit; transition: all 0.15s;">
          <img src="assets/icon-80.svg" alt="" style="width:20px;height:20px;vertical-align:middle;margin-right:6px;border-radius:4px;"> Try Kuro
        </button>
      </div>
    `;
    appBody.style.display = "flex";

    // Hide header
    const header = document.querySelector("header.ms-welcome__header");
    if (header) {
      header.style.display = "none";
      header.style.height = "0";
      header.style.overflow = "hidden";
      header.style.margin = "0";
      header.style.padding = "0";
    }

    // Open chat button
    const openChatBtn = document.getElementById("open-chat-btn");
    if (openChatBtn) {
      openChatBtn.onclick = () => {
        appBody.style.display = "none";
        chatContainer.style.display = "flex";
      };
    }

    // Ref to DOM elements
    const endpointInput = document.getElementById("agent-endpoint");
    const modelInput = document.getElementById("agent-model");
    const chatInput = document.getElementById("chat-input");
    const sendButton = document.getElementById("chat-send");
    const stopButton = document.getElementById("chat-stop");
    const imageInput = document.getElementById("chat-image-input");
    const imageUploadBtn = document.getElementById("chat-image-upload");

    console.log("[agentChat] endpointInput:", endpointInput);
    console.log("[agentChat] chatContainer:", chatContainer);

    // ─── Model Discovery ─────────────────────────────────────────────────────

    const modelStatus = document.getElementById("model-status");
    let modelFetchAbort = null;

    function showModelLoading() {
      if (modelStatus) {
        modelStatus.innerHTML =
          '<span class="model-loading"><span class="spinner"></span> Discovering models...</span>';
      }
    }

    function showModelCount(count) {
      if (modelStatus) {
        modelStatus.innerHTML = `<span class="model-count">${count} models discovered</span>`;
      }
    }

    function showModelError() {
      if (modelStatus) {
        modelStatus.innerHTML = `<span class="model-count" style="color:var(--chat-error-text)">Manual entry supported</span>`;
      }
    }

    function clearModelStatus() {
      if (modelStatus) {
        modelStatus.innerHTML = "";
      }
    }

    async function discoverModels() {
      if (modelFetchAbort) {
        modelFetchAbort.abort();
      }

      const currentEndpoint = getConfig().endpoint;
      const currentApiKey = getConfig().apiKey;

      // Don't fetch if no API key or using default placeholder
      if (!currentApiKey || currentApiKey.includes("PLACEHOLDER")) {
        clearModelStatus();
        return;
      }

      modelFetchAbort = new AbortController();
      showModelLoading();

      try {
        const models = await fetchModels(currentEndpoint, currentApiKey);

        if (modelFetchAbort.signal.aborted) return;

        if (models.length > 0) {
          // Build datalist from discovered models only
          const datalist = document.getElementById("model-suggestions");
          if (datalist) {
            datalist.innerHTML = "";
            models.forEach((m) => {
              const opt = document.createElement("option");
              opt.value = m;
              datalist.appendChild(opt);
            });

            showModelCount(models.length);
          }
        } else {
          showModelError();
        }
      } catch (err) {
        if (modelFetchAbort.signal.aborted) return;
        console.warn("[agentChat] Model discovery failed:", err.message);
        showModelError();
      }
    }

    // Debounce model discovery
    let modelDiscoverTimer = null;
    function scheduleModelDiscover() {
      if (modelDiscoverTimer) clearTimeout(modelDiscoverTimer);
      modelDiscoverTimer = setTimeout(discoverModels, 500);
    }

    if (endpointInput) {
      endpointInput.addEventListener("change", () => {
        setAiConfig({ endpoint: endpointInput.value.trim() });
        scheduleModelDiscover();
      });
    }

    if (modelInput) {
      modelInput.addEventListener("change", () => {
        setAiConfig({ model: modelInput.value.trim() });
      });
    }

    const apiKeyInput = document.getElementById("agent-apikey");
    if (apiKeyInput) {
      apiKeyInput.addEventListener("change", () => {
        setAiConfig({ apiKey: apiKeyInput.value.trim() });
        scheduleModelDiscover();
      });
    }

    const telemetryToggle = document.getElementById("telemetry-toggle");
    const telemetryStatus = document.getElementById("telemetry-status");
    if (telemetryToggle) {
      telemetryToggle.addEventListener("change", () => {
        const enabled = telemetryToggle.checked;
        localStorage.setItem("telemetryEnabled", enabled ? "true" : "false");
        telemetryStatus.textContent = enabled ? "On" : "Off";
        telemetryStatus.style.color = enabled ? "var(--chat-text-muted)" : "var(--chat-error-text)";
      });
    }

    // Mode toggle buttons
    const interactiveModeBtn = document.getElementById("mode-interactive");
    const explainModeBtn = document.getElementById("mode-explain");
    if (interactiveModeBtn) {
      interactiveModeBtn.addEventListener("click", () => switchMode("interactive"));
    }
    if (explainModeBtn) {
      explainModeBtn.addEventListener("click", () => switchMode("explain"));
    }

    if (chatInput) {
      chatInput.addEventListener("input", () => {
        chatInput.style.height = "auto";
        chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + "px";
        sendButton.disabled = chatInput.value.trim() === "" && !isExecuting;
      });
    }

    if (chatInput) {
      chatInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          handleSend();
        } else if (e.key === "Tab") {
          e.preventDefault();
          const nextMode = currentMode === "interactive" ? "explain" : "interactive";
          switchMode(nextMode);
          chatInput.placeholder =
            currentMode === "interactive"
              ? `Ready to modify – describe what to do in ${currentHost === "powerpoint" ? "PowerPoint" : "Excel"}...`
              : "Read-only – ask about your content (Tab to switch)...";
          clearTimeout(chatInput._placeholderTimeout);
          chatInput._placeholderTimeout = setTimeout(() => {
            chatInput.placeholder = `Ask Kuro anything about your ${currentHost === "powerpoint" ? "slides" : "spreadsheet"}...`;
          }, 2000);
        }
      });
    }

    // Smart paste detection
    if (chatInput) {
      chatInput.addEventListener("paste", (e) => {
        const pastedText = (e.clipboardData || window.clipboardData).getData("text");
        const parsed = parseTSV(pastedText);
        if (parsed && parsed.rows >= 1 && parsed.cols >= 1) {
          e.preventDefault();
          showPastePreview(parsed.matrix, parsed.rows, parsed.cols, {
            appendMessage,
            scrollToBottom: scrollChatToBottom,
            onPasteWritten: (rows, cols) => {
              conversationHistory.push({
                role: "assistant",
                content: `Wrote ${rows}×${cols} data to active ${currentHost === "powerpoint" ? "slide" : "worksheet"}.`,
              });
            },
          });
          chatInput.value = "";
          chatInput.style.height = "auto";
          sendButton.disabled = true;
        }
      });
    }

    if (sendButton) {
      sendButton.addEventListener("click", handleSend);
    }

    if (stopButton) {
      stopButton.addEventListener("click", () => {
        stopAgent();
        stopButton.style.display = "none";
        sendButton.disabled = false;
        chatInput.focus();
      });
    }

    if (imageInput) {
      imageInput.addEventListener("change", (e) => {
        const file = e.target.files && e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function (event) {
          pendingImageBase64 = event.target.result;
          showImagePreview(pendingImageBase64);
          imageUploadBtn.style.display = "inline-block";
          console.log("[agentChat] Image attached:", file.name);
        };
        reader.readAsDataURL(file);
        imageInput.value = "";
      });
    }

    if (imageUploadBtn) {
      imageUploadBtn.addEventListener("click", () => {
        if (pendingImageBase64) {
          pendingImageBase64 = null;
          hideImagePreview();
          imageUploadBtn.style.display = "none";
          chatInput.focus();
        } else {
          imageInput.click();
        }
      });
    }

    const clearButton = document.getElementById("chat-clear");
    if (clearButton) {
      clearButton.addEventListener("click", () => {
        if (isExecuting) {
          stopAgent();
          document.getElementById("chat-stop").style.display = "none";
          sendButton.disabled = false;
        }
        clearChat();
      });
    }

    // ─── Config Toggle & Auto-Deflate ────────────────────────────────────────

    const configBar = document.getElementById("agent-config");
    const configToggle = document.getElementById("config-toggle");
    let configExpanded = localStorage.getItem("agentConfigExpanded");
    if (configExpanded === null) configExpanded = false; // Default to collapsed
    configExpanded = configExpanded === "true";
    let autoDeflateTimer = null;
    const AUTO_DEFDELAY = 15000; // 15 seconds

    function setConfigState(expanded) {
      configExpanded = expanded;
      localStorage.setItem("agentConfigExpanded", expanded ? "true" : "false");
      if (configBar) {
        configBar.classList.toggle("collapsed", !expanded);
        configBar.classList.toggle("expanded", expanded);
      }
      const toggleText = document.getElementById("toggle-text");
      if (toggleText) {
        toggleText.textContent = expanded ? "Collapse" : "Settings";
      }
      resetAutoDeflate();
    }

    function resetAutoDeflate() {
      if (autoDeflateTimer) clearTimeout(autoDeflateTimer);
      if (configExpanded && isExecuting) return; // Don't auto-deflate during execution
      if (configExpanded) {
        autoDeflateTimer = setTimeout(() => {
          if (configExpanded && !isExecuting) {
            setConfigState(false);
          }
        }, AUTO_DEFDELAY);
      }
    }

    if (configToggle) {
      configToggle.addEventListener("click", () => {
        setConfigState(!configExpanded);
      });
    }

    // Pause auto-deflate on any config interaction
    const configInputs = configBar ? configBar.querySelectorAll("input, button") : [];
    configInputs.forEach((input) => {
      input.addEventListener("focus", () => {
        if (!configExpanded) setConfigState(true);
        resetAutoDeflate();
      });
      input.addEventListener("blur", () => resetAutoDeflate());
    });

    // Pause auto-deflate during execution, resume after
    const originalIsExecuting = { value: false };
    Object.defineProperty(window, "isExecuting", {
      get: () => originalIsExecuting.value,
      set: (val) => {
        originalIsExecuting.value = val;
        if (val) {
          if (autoDeflateTimer) clearTimeout(autoDeflateTimer);
        } else {
          resetAutoDeflate();
        }
      },
      configurable: true,
    });

    // Initialize config state
    setConfigState(configExpanded);

    // Initial model discovery if endpoint+apiKey already configured
    if (getConfig().apiKey && !getConfig().apiKey.includes("PLACEHOLDER")) {
      discoverModels();
    }

    if (chatInput) {
      chatInput.focus();
    }

    window.addEventListener("beforeunload", () => {
      flushTelemetry(false);
    });

    const aiConf = getConfig();
    initTelemetry(aiConf.model, DEFAULTS.telemetryEndpoint);

    console.log("[agentChat] initAgentChat complete, host:", currentHost);
  } catch (err) {
    console.error("[agentChat] Error during initialization:", err);
  }
}

// ─── Main Send Handler ──────────────────────────────────────────────────────

async function handleSend() {
  await loadHostModules();
  const chatInput = document.getElementById("chat-input");
  const message = chatInput.value.trim();
  if (!message && !pendingImageBase64) return;
  if (isExecuting) return;

  chatInput.value = "";
  chatInput.style.height = "auto";

  if (pendingImageBase64) {
    appendUserMessageWithImage(message, pendingImageBase64);
    conversationHistory.push({
      role: "user",
      content: [
        { type: "text", text: message },
        { type: "image_url", image_url: { url: pendingImageBase64 } },
      ],
    });
    console.log("[agentChat] Sending with image to AI");
  } else {
    appendMessage("user", message);
    conversationHistory.push({ role: "user", content: message });
  }

  originalQuery = message;
  feedbackStepData = null;
  feedbackTotalSteps = 0;
  feedbackTotalTiming = 0;

  // Capture before snapshot for telemetry
  try {
    const beforeContext = await getHostMetadataFn()();
    if (beforeContext) {
      trackSnapshot("before", beforeContext);
    }
  } catch {
    // Non-critical
  }

  isExecuting = true;
  isStopped = false;
  const sendButton = document.getElementById("chat-send");
  const stopButton = document.getElementById("chat-stop");
  if (sendButton) sendButton.disabled = true;

  // Declare stream before the try block so catch/finally can access it
  let stream = null;

  try {
    if (stopButton) stopButton.style.display = "inline-block";

    // Phase 1: Planning
    console.log("[plan] === Starting planning phase ===");
    let sheetContext = await getHostContextFn()();
    console.log("[plan] Context:", sheetContext);
    let systemPrompt = buildSystemPrompt(currentHost, currentMode, sheetContext);

    const apiMessages = [
      { role: "system", content: systemPrompt },
      ...(await summarizeConversationHistory(conversationHistory, streamFromAI)),
    ];
    console.log(
      "[plan] Sending to AI — system prompt length:",
      systemPrompt.length,
      "msg count:",
      apiMessages.length
    );

    stream = appendStreamingMessage();
    let planResponse = "";

    try {
      planResponse = await streamFromAI(
        apiMessages,
        () => isStopped,
        undefined,
        (text, isComplete) => {
          if (isComplete) {
            stream.complete();
          } else {
            stream.update(text);
          }
        }
      );
    } catch (error) {
      stream?.element?.remove();
      appendMessage("agent", `Error: ${error.message || String(error)}`);
      return;
    }

    if (!planResponse && stream?.element && !stream.element.querySelector(".agent-bubble").innerHTML) {
      stream.element.remove();
      if (isStopped) {
        appendMessage("agent", "⏹ Execution stopped.");
      } else {
        appendMessage("agent", "No response from agent.");
      }
      trackStep({
        stepNumber: 0,
        userPrompt: originalQuery,
        plan: "",
        operations: [],
        results: [],
        errors: isStopped
          ? ["Execution stopped by user"]
          : ["AI returned null or empty response during planning"],
        success: false,
        timingMs: 0,
        aiResponse: planResponse || "(null)",
        verification: "",
        conversationHistory: apiMessages.map((m) => ({
          role: m.role,
          content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
        })),
        sheetContext,
        systemPrompt,
      });
      return;
    }

    const ops = extractOperations(planResponse);
    console.log("[plan] Extracted operations:", ops);

    // If JSON was found but unparseable, retry once asking the model to fix it
    let finalOps = ops;
    if (
      ops &&
      ops.operations.length === 0 &&
      ops.text &&
      /```json/.test(ops.text)
    ) {
      console.log("[plan] JSON block found but unparseable — retrying with fix request...");
      const fixMessages = [
        ...apiMessages,
        { role: "assistant", content: planResponse },
        {
          role: "user",
          content:
            "Your JSON response had a syntax error and could not be parsed. " +
            "Please fix the JSON and respond again with a valid ```json block. " +
            "Common issues: missing commas between array elements, unclosed strings, unescaped quotes inside strings. " +
            "Respond with ONLY the corrected JSON block, no explanation.",
        },
      ];

      try {
        const fixResponse = await streamFromAI(
          fixMessages,
          () => isStopped,
          undefined,
          (text, isComplete) => {
            if (isComplete) {
              stream.complete();
            } else {
              stream.update(text);
            }
          }
        );
        if (fixResponse) {
          finalOps = extractOperations(fixResponse);
          if (finalOps && finalOps.operations.length > 0) {
            console.log("[plan] Retry succeeded —", finalOps.operations.length, "operations");
            planResponse = fixResponse;
          } else {
            console.warn("[plan] Retry also failed to produce valid operations");
          }
        }
      } catch (retryErr) {
        console.warn("[plan] Retry request failed:", retryErr.message);
      }
    }

    if (!finalOps) {
      stream.element.remove();
      appendMessage("agent", "No response from agent.");
      conversationHistory.push({ role: "assistant", content: planResponse });
      trackStep({
        stepNumber: 0,
        userPrompt: originalQuery,
        plan: "",
        operations: [],
        results: [],
        errors: ["AI response could not be parsed into structured operations"],
        success: false,
        timingMs: 0,
        aiResponse: planResponse,
        verification: "",
        conversationHistory: apiMessages.map((m) => ({
          role: m.role,
          content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
        })),
        sheetContext,
        systemPrompt,
      });
      return;
    }

    // Show plan
    if (finalOps.plan) {
      appendPlanMessage(finalOps.plan);
    }

    // If no operations — AI is chatting or requesting mode switch
    if (finalOps.operations.length === 0) {
      stream.element.remove();
      if (finalOps._switchMode === "interactive" && currentMode === "explain") {
        const switchMsg = finalOps._switchMessage || "Switch to interactive mode for modifications?";
        const msgEl = appendMessage("agent", planResponse || switchMsg);
        const switchBtn = document.createElement("button");
        switchBtn.className = "agent-switch-mode-btn";
        switchBtn.textContent = "Switch to Interactive Mode";
        switchBtn.style.marginTop = "8px";
        switchBtn.addEventListener("click", () => {
          switchMode("interactive");
          switchBtn.remove();
        });
        msgEl.appendChild(switchBtn);
      } else {
        appendMessage("agent", planResponse || "No operations generated.");
      }
      conversationHistory.push({ role: "assistant", content: planResponse });
      trackStep({
        stepNumber: 0,
        userPrompt: originalQuery,
        plan: finalOps.plan || "",
        operations: [],
        results: [],
        errors: ["AI returned 0 operations — pure chat mode"],
        success: false,
        timingMs: 0,
        aiResponse: planResponse,
        verification: "",
        conversationHistory: apiMessages.map((m) => ({
          role: m.role,
          content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
        })),
        sheetContext,
        systemPrompt,
      });
      return;
    }

    // Validate operations
    const registry =
      currentHost === "powerpoint"
        ? hostModules.ops.PPT_OPERATION_REGISTRY
        : EXCEL_OPERATION_REGISTRY;
    const { valid: validOps, errors: validationErrors } = validateOperations(
      finalOps.operations,
      registry,
      currentMode
    );
    console.log(
      "[plan] Validation — valid:",
      validOps.length,
      "errors:",
      validationErrors.length,
      validationErrors
    );

    if (validationErrors.length > 0) {
      stream.element.remove();
      appendMessage("agent", `⚠ Validation errors: ${validationErrors.join("; ")}`);
      conversationHistory.push({ role: "assistant", content: planResponse });
      trackStep({
        stepNumber: 0,
        userPrompt: originalQuery,
        plan: finalOps.plan || "",
        operations: finalOps.operations,
        results: [],
        errors: validationErrors,
        success: false,
        timingMs: 0,
        aiResponse: planResponse,
        verification: "",
        conversationHistory: apiMessages.map((m) => ({
          role: m.role,
          content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
        })),
        sheetContext,
        systemPrompt,
      });
      return;
    }

    // Phase 2: Execute operations loop
    await executeOperationsLoop(ops, sheetContext, systemPrompt);
  } catch (error) {
    console.error("[agentChat] handleSend error:", error);
    stream?.element?.remove();
    appendMessage("agent", `Error: ${error.message || String(error)}`);
  } finally {
    isExecuting = false;
    isStopped = false;
    if (sendButton) sendButton.disabled = false;
    if (stopButton) stopButton.style.display = "none";
    chatInput.focus();
  }
}
