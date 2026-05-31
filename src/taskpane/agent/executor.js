/* global console, document, Excel, PowerPoint */

/* eslint-disable no-undef */

/* Execution engine — handleSend, executeOperationsLoop, retry/revert, continue */

import { streamFromAI, getConfig } from "./ai.js";
import { buildSystemPrompt } from "./prompts.js";
import {
  extractOperations,
  validateOperations,
  summarizeConversationHistory,
  isCompletionClaim,
} from "./parser.js";
import { EXCEL_OPERATION_REGISTRY, executeOperationWithTracking } from "./operations.js";
import {
  appendMessage,
  appendPlanMessage,
  appendStepMessage,
  appendExecutionResult,
  scrollChatToBottom,
  setExecutingState,
  appendStreamingMessage,
  finalizeStepGroup,
} from "../ui/chat.js";
import {
  trackStep,
  trackSnapshot,
  submitFeedback,
  trackPhase,
  trackConsecutiveFailure,
} from "../telemetry.js";
import { S, MAX_CONSECUTIVE_FAILURES, MAX_NO_OPS_RETRIES } from "../store.js";
import {
  loadHostModules,
  getHostContextFn,
  getHostFullContextFn,
  getHostMetadataFn,
  getHostCaptureSnapshotFn,
  getHostVerifyFn,
  getHostExtractNamesFn,
  getHostModules,
} from "./host.js";
import { escapeHtml } from "../ui/chat.js";
import { buildCompletionSummary, appendFeedbackButtons } from "../ui/config.js";

// ─── Retry / Revert ──────────────────────────────────────────────────────────

async function handleRetry(stepIndex, onAcknowledge, onImprove) {
  if (stepIndex < 0 || stepIndex >= S.stepStack.length) return;
  const step = S.stepStack[stepIndex];

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

  S.conversationHistory.push({ role: "user", content: retryContent });
  S.stepStack.splice(stepIndex, S.stepStack.length - stepIndex);

  if (step.resultEl && step.resultEl.parentElement) {
    step.resultEl.parentElement.removeChild(step.resultEl);
  }

  await continueExecutionFromAI(onAcknowledge, onImprove);
}

async function handleRevert(stepIndex, onAcknowledge, onImprove) {
  if (stepIndex < 0 || stepIndex >= S.stepStack.length) return;
  const step = S.stepStack[stepIndex];

  const opSummary = step.operations
    .map((o) => `${o.operation}(${o.sheet}, ${o.range || o.name || ""})`)
    .join(", ");

  const revertContent =
    `I want to undo/revert the action from Step ${step.stepNumber}. ` +
    `Operations that were executed: ${opSummary}\n` +
    `Plan: ${step.plan}\n` +
    `Please undo these changes. If you have access to the previous state, restore it. ` +
    `Otherwise, reverse each operation manually.`;

  S.conversationHistory.push({ role: "user", content: revertContent });
  S.stepStack.splice(stepIndex, S.stepStack.length - stepIndex);

  if (step.resultEl && step.resultEl.parentElement) {
    step.resultEl.parentElement.removeChild(step.resultEl);
  }

  await continueExecutionFromAI(onAcknowledge, onImprove);
}

// ─── Continue Execution (after retry/revert) ────────────────────────────────

async function continueExecutionFromAI(onAcknowledge, onImprove) {
  await loadHostModules();
  const sheetContext = await getHostContextFn()();
  const systemPrompt = buildSystemPrompt(S.currentHost, S.currentMode, sheetContext);

  const nextApiMessages = [
    { role: "system", content: systemPrompt },
    ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
  ];

  const nextResponse = await streamFromAI(nextApiMessages, () => S.isStopped);
  if (!nextResponse || S.isStopped) return;

  S.conversationHistory.push({ role: "assistant", content: nextResponse });
  const nextOps = extractOperations(nextResponse);
  if (!nextOps || nextOps.operations.length === 0) {
    if (isCompletionClaim(nextResponse)) {
      console.log("[validate] AI claims complete (retry path) — validating...");
      appendMessage("agent", "🔍 Verifying completion...");

      const fullContext = await getHostFullContextFn()();
      const hostLabel = S.currentHost === "powerpoint" ? "presentation" : "workbook";
      const validationMessages = [
        {
          role: "system",
          content: `You are a ${S.currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task validator. Your ONLY job is to check whether the agent completed the user's request. You are NOT an agent — you do NOT generate operations.

USER REQUEST: "${S.originalQuery}"

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
        ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
      ];

      const validationResult = await streamFromAI(validationMessages, () => S.isStopped);
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
          onAcknowledge();
          return;
        } else {
          const issuesText =
            parsedResult.issues.length > 0
              ? parsedResult.issues.join(", ")
              : "Validation inconclusive";
          appendMessage("agent", `⚠ Issues: ${issuesText}. Asking agent to fix...`);

          const fixMessages = [
            { role: "system", content: systemPrompt },
            ...S.conversationHistory,
            {
              role: "user",
              content: `⚠ Validation failed: ${parsedResult.issues.join("\n")}. Please fix with structured operations.`,
            },
          ];

          const fixResponse = await streamFromAI(fixMessages, () => S.isStopped);
          if (!fixResponse || S.isStopped) return;

          S.conversationHistory.push({ role: "assistant", content: fixResponse });
          const fixOps = extractOperations(fixResponse);
          if (fixOps && fixOps.operations.length > 0) {
            await executeOperationsLoop(
              fixOps,
              sheetContext,
              systemPrompt,
              onAcknowledge,
              onImprove
            );
            return;
          } else {
            appendMessage("agent", fixResponse.substring(0, 300));
            return;
          }
        }
      } else {
        appendMessage("agent", "⚠ Validation skipped.");
        onAcknowledge();
        return;
      }
    }
    appendMessage("agent", "⚠ No operations found. Asking to continue...");
    const retryFeedback =
      "Your last response did not contain structured operations. Please provide the next step as structured operations, or say 'All steps complete.' if done.";
    S.conversationHistory.push({ role: "user", content: retryFeedback });
    const retryResponse = await streamFromAI(
      [
        { role: "system", content: systemPrompt },
        ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
      ],
      () => S.isStopped
    );
    if (!retryResponse || S.isStopped) return;
    S.conversationHistory.push({ role: "assistant", content: retryResponse });
    const retryOps = extractOperations(retryResponse);
    if (!retryOps || retryOps.operations.length === 0) {
      appendMessage("agent", "All steps complete.");
      S.conversationHistory = [];
      S.stepStack = [];
      S.originalQuery = "";
      return;
    }
    await executeOperationsLoop(retryOps, sheetContext, systemPrompt, onAcknowledge, onImprove);
    return;
  }
  await executeOperationsLoop(nextOps, sheetContext, systemPrompt, onAcknowledge, onImprove);
}

// ─── Operations Execution Loop ───────────────────────────────────────────────

async function executeOperationsLoop(
  initialOps,
  initialContext,
  systemPrompt,
  onAcknowledge,
  onImprove
) {
  const plan = initialOps?.plan || "";
  const opsArray = Array.isArray(initialOps) ? initialOps : initialOps?.operations || [];
  let validOps = [...opsArray];
  let sheetContext = initialContext;
  let stepNumber = 1;

  while (validOps.length > 0 && !S.isStopped) {
    const stepStartTime = Date.now();
    appendStepMessage(stepNumber, plan, validOps);
    stepNumber++;

    const snapshotBefore = await getHostCaptureSnapshotFn()(validOps);

    const results = [];
    let hadError = false;

    for (const op of validOps) {
      try {
        const runFn = S.currentHost === "powerpoint" ? PowerPoint.run : Excel.run;
        const result = await runFn(async (context) => {
          return executeOperationWithTracking(context, op);
        });
        results.push(result);
      } catch (error) {
        results.push(`✗ ${op.operation} failed: ${error.message}`);
        hadError = true;

        // Record error in telemetry
        try {
          const { recordOpResult } = await import("../telemetry.js");
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
      _onRetry: (idx) => handleRetry(idx, onAcknowledge, onImprove),
      _onRevert: (idx) => handleRevert(idx, onAcknowledge, onImprove),
    });

    // Feed execution result back to AI
    let feedback;
    let verification = "";
    if (hadError) {
      S.consecutiveFailures++;
      // Track consecutive failure for telemetry
      try {
        trackConsecutiveFailure("execution", S.consecutiveFailures);
      } catch (e) {
        // Non-critical
      }

      if (S.consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
        appendMessage(
          "agent",
          `❌ ${MAX_CONSECUTIVE_FAILURES} consecutive failures. Stopping to avoid wasted API calls.`
        );
        appendMessage("agent", "Please try again or describe the issue more specifically.");
        S.conversationHistory = [];
        S.stepStack = [];
        S.originalQuery = "";
        break;
      }
      feedback = `Some operations failed: ${resultText}. Adapt and provide corrected operations. (Failure ${S.consecutiveFailures}/${MAX_CONSECUTIVE_FAILURES})`;
      console.log(`[step:${stepNumber}] Feedback sent (error):`, feedback);
      S.conversationHistory.push({ role: "user", content: feedback });
    } else {
      S.consecutiveFailures = 0;
      const opField = S.currentHost === "powerpoint" ? "slide" : "sheet";
      const actualNames = getHostExtractNamesFn()(results, validOps);
      verification = await getHostVerifyFn()(validOps, actualNames);
      const opSummary = validOps
        .map((o) => `${o.operation}(${o[opField] || o.name || ""}, ${o.range || o.shape || ""})`)
        .join(", ");
      const resultsSection = results.length > 0 ? `\n\nOperation results:\n${resultText}` : "";
      const verificationSection = verification ? `\n\nVerification:\n${verification}` : "";
      feedback = `Step executed. Operations: ${opSummary}${resultsSection}${verificationSection}\n\nCurrent ${S.currentHost === "powerpoint" ? "slide" : "sheet"} state is in your system prompt. Provide the next step as structured operations, or say "All steps complete." if done.`;
      console.log(`[step:${stepNumber}] Feedback sent (success):`, feedback);
      S.conversationHistory.push({ role: "user", content: feedback });
    }

    S.stepStack.push({
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
      userPrompt: S.originalQuery,
      plan,
      operations: validOps,
      results: results,
      errors: hadError ? [resultText] : [],
      success: !hadError,
      timingMs: Date.now() - stepStartTime,
      aiResponse: nextResponse || "",
      verification: verification || "",
      conversationHistory: [...S.conversationHistory],
      sheetContext,
      systemPrompt,
      aiModel: getConfig().model,
      aiEndpoint: getConfig().endpoint,
      phase: "execution",
      consecutiveFailures: S.consecutiveFailures,
    });

    // Track execution phase
    try {
      trackPhase("execution", !hadError, Date.now() - stepStartTime);
    } catch (e) {
      // Non-critical
    }

    const stepTiming = Date.now() - stepStartTime;
    S.feedbackTotalSteps++;
    S.feedbackTotalTiming += stepTiming;

    // Read updated context
    sheetContext = await getHostContextFn()();
    console.log(`[step:${stepNumber}] Updated context:`, sheetContext);
    systemPrompt = buildSystemPrompt(S.currentHost, S.currentMode, sheetContext);

    const nextApiMessages = [
      { role: "system", content: systemPrompt },
      ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
    ];
    console.log(
      `[step:${stepNumber}] Sending to AI — msg count:`,
      nextApiMessages.length,
      "system prompt length:",
      systemPrompt.length
    );

    const nextResponse = await streamFromAI(nextApiMessages, () => S.isStopped);
    console.log(
      `[step:${stepNumber}] AI response received (length:`,
      nextResponse ? nextResponse.length : 0,
      `):\n`,
      nextResponse
    );

    if (!nextResponse || S.isStopped) {
      if (S.isStopped) {
        appendMessage("agent", "⏹ Execution stopped.");
      } else {
        appendMessage("agent", "All steps complete.");
      }
      if (!S.isStopped) {
        S.conversationHistory = [];
        S.stepStack = [];
        S.originalQuery = "";
      }
      break;
    }

    S.conversationHistory.push({ role: "assistant", content: nextResponse });
    const nextOps = extractOperations(nextResponse);
    console.log(`[step:${stepNumber}] Extracted ops:`, nextOps);
    if (!nextOps || nextOps.operations.length === 0) {
      if (isCompletionClaim(nextResponse)) {
        console.log(
          "[validate] AI claims complete — validating against full",
          S.currentHost,
          "state..."
        );
        appendMessage("agent", `🔍 Verifying completion against full ${S.currentHost} state...`);

        // Track validation phase start
        const validationStartTime = Date.now();
        try {
          trackPhase("validation", true, 0);
        } catch (e) {
          // Non-critical
        }

        const fullContext = await getHostFullContextFn()();
        console.log("[validate] Full context:\n", fullContext);

        const hostLabel = S.currentHost === "powerpoint" ? "presentation" : "workbook";
        const opField = S.currentHost === "powerpoint" ? "slide" : "sheet";
        const opsList = validOps
          .map(
            (o) =>
              `- ${o.operation}${o[opField] ? ` on ${o[opField]}` : ""}${o.range ? ` (${o.range})` : ""}`
          )
          .join("\n");
        const validationMessages = [
          {
            role: "system",
            content: `You are a ${S.currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task validator. Your ONLY job is to check whether the agent completed the user's request. You are NOT an agent — you do NOT generate operations.

USER REQUEST: "${S.originalQuery}"

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
          ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
        ];

        const validationResult = await streamFromAI(validationMessages, () => S.isStopped);
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
            const validationTiming = Date.now() - validationStartTime;

            // Track validation phase success
            try {
              trackPhase("validation", true, validationTiming);
            } catch (e) {
              // Non-critical
            }

            try {
              const afterContext = await getHostMetadataFn()();
              if (afterContext) {
                trackSnapshot("after", afterContext);
              }
            } catch {
              // Non-critical
            }

            const opField = S.currentHost === "powerpoint" ? "slide" : "sheet";
            S.feedbackStepData = {
              userPrompt: S.originalQuery,
              plan,
              operations: validOps,
              results: results,
              errors: [],
              success: true,
              timingMs: S.feedbackTotalTiming,
            };

            trackStep({
              stepNumber: S.feedbackTotalSteps,
              userPrompt: S.originalQuery,
              plan,
              operations: validOps,
              results: results,
              errors: [],
              success: true,
              timingMs: S.feedbackTotalTiming,
              aiResponse: nextResponse || "",
              verification: verification || "",
              conversationHistory: [...S.conversationHistory],
              sheetContext,
              systemPrompt,
            });

            const taskStats = {
              totalSteps: S.feedbackTotalSteps,
              totalTiming: S.feedbackTotalTiming,
              finalSuccess: true,
            };

            // Generate summary
            const summaryMessages = [
              {
                role: "system",
                content: `You are a ${S.currentHost === "powerpoint" ? "PowerPoint" : "Excel"} task summarizer. Based on the operations performed and the final ${S.currentHost === "powerpoint" ? "presentation" : "workbook"} state, provide a clear, user-friendly summary of what was accomplished.

Respond with a JSON object:
{
  "title": "Brief title of what was done",
  "changes": ["list of specific changes made, one per line"],
  "details": "More detailed explanation of the changes and their impact"
}`,
              },
              {
                role: "user",
                content: `Original request: "${S.originalQuery}"

Operations performed:
${validOps.map((op) => `- ${op.operation} on ${op[opField] || op.name}${op.range ? ` (${op.range})` : ""}`).join("\n")}

Final ${S.currentHost === "powerpoint" ? "presentation" : "workbook"} state:
${fullContext}

Provide a clear summary of what was done.`,
              },
            ];

            let summaryText = "";
            try {
              const summaryResponse = await streamFromAI(summaryMessages, () => S.isStopped);
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
                  onAcknowledge();
                },
                onBad: () => {
                  submitFeedback("bad", "", taskStats);
                  console.log("[telemetry] User rated: Bad");
                  onAcknowledge();
                },
                onImprove: (comment) => {
                  onImprove(comment, taskStats);
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
            const validationFailTiming = Date.now() - validationStartTime;

            // Track validation phase failure
            try {
              trackPhase("validation", false, validationFailTiming);
            } catch (e) {
              // Non-critical
            }

            const fixMessages = [
              { role: "system", content: systemPrompt },
              ...S.conversationHistory,
              {
                role: "user",
                content: `⚠ Validation failed. The following issues were found:\n${parsedResult.issues.join("\n")}\n\nPlease fix these issues with structured operations, or if everything is actually correct, explain why in plain text.`,
              },
            ];

            // Track fix phase start
            try {
              trackPhase("fix", true, 0);
            } catch (e) {
              // Non-critical
            }

            const fixResponse = await streamFromAI(fixMessages, () => S.isStopped);
            console.log(
              "[validate] Fix response received:",
              fixResponse ? `${fixResponse.length} chars` : "null/empty"
            );
            if (!fixResponse || S.isStopped) {
              // Track fix phase failure
              try {
                trackPhase("fix", false, Date.now() - validationStartTime);
              } catch (e) {
                // Non-critical
              }

              if (S.isStopped) {
                appendMessage("agent", "⏹ Execution stopped.");
              } else {
                appendMessage(
                  "agent",
                  "⚠ Agent did not respond to the fix request. Try again manually."
                );
              }
              break;
            }

            S.conversationHistory.push({ role: "assistant", content: fixResponse });
            const fixOps = extractOperations(fixResponse);
            console.log(
              "[validate] Fix ops extracted:",
              fixOps ? fixOps.operations.length : 0,
              "operations"
            );
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
                userPrompt: S.originalQuery,
                plan,
                operations: validOps,
                results: results,
                errors: ["Validation fix AI responded with plain text instead of operations"],
                success: false,
                timingMs: 0,
                aiResponse: fixResponse || "(null)",
                verification: verification || "",
                conversationHistory: [...S.conversationHistory],
                sheetContext,
                systemPrompt,
              });
              break;
            }
          }
        } else {
          appendMessage("agent", "⚠ Validation skipped — no response from validator.");
          onAcknowledge();
          return;
        }
      }

      // No operations — ask for clarification
      S.noOpsRetries++;
      if (S.noOpsRetries > MAX_NO_OPS_RETRIES) {
        appendMessage(
          "agent",
          "⚠ I've tried multiple times but can't determine the next steps. Would you like to describe what to do next?"
        );
        S.isExecuting = false;
        const sendBtn = document.getElementById("chat-send");
        if (sendBtn) sendBtn.disabled = false;
        return;
      }
      appendMessage("agent", "⚠ No operations found. Asking to continue...");
      const retryFeedback =
        "Your last response did not contain structured operations. Please provide the next step as structured operations, or say 'All steps complete.' if done.";
      S.conversationHistory.push({ role: "user", content: retryFeedback });

      const retryResponse = await streamFromAI(
        [
          { role: "system", content: systemPrompt },
          ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
        ],
        () => S.isStopped
      );

      if (!retryResponse) break;

      S.conversationHistory.push({ role: "assistant", content: retryResponse });
      const retryOps = extractOperations(retryResponse);
      if (!retryOps || retryOps.operations.length === 0) {
        appendMessage("agent", "⚠ Agent could not generate operations. Stopping.");
        break;
      }

      validOps.length = 0;
      validOps.push(...retryOps.operations);
      S.noOpsRetries = 0;
    } else {
      validOps.length = 0;
      validOps.push(...nextOps.operations);
      S.noOpsRetries = 0;
    }
  }

  finalizeStepGroup();
}

// ─── Improve Flow ────────────────────────────────────────────────────────────

async function handleImprove(comment, taskStats, onAcknowledge) {
  const messagesContainer = document.getElementById("chat-messages");
  if (!messagesContainer || !S.feedbackStepData) return;

  const improveArea = messagesContainer.querySelector(".agent-improve-area");
  if (improveArea) improveArea.style.display = "none";

  const buttons = messagesContainer.querySelectorAll(".agent-feedback-btn");
  buttons.forEach((btn) => (btn.disabled = true));

  appendMessage("agent", `🔧 Applying improvements: "${comment}"`);

  // Track improvement phase start
  const improveStartTime = Date.now();
  try {
    trackPhase("improvement", true, 0);
  } catch (e) {
    // Non-critical
  }

  await loadHostModules();
  const sheetContext = await getHostContextFn()();
  const systemPrompt = buildSystemPrompt(S.currentHost, S.currentMode, sheetContext);

  const opField = S.currentHost === "powerpoint" ? "slide" : "sheet";
  const feedbackMessages = [
    { role: "system", content: systemPrompt },
    ...S.conversationHistory,
    {
      role: "user",
      content: `The user wants to improve the previous result.

Original request: "${S.originalQuery}"

Previous operations executed:
${S.feedbackStepData.operations.map((o) => `- ${o.operation}${o[opField] ? ` on ${o[opField]}` : ""}${o.range ? ` (${o.range})` : ""}`).join("\n")}

User feedback: "${comment}"

Please generate new structured operations to incorporate this improvement. Output inside a JSON code block.`,
    },
  ];

  const improveResponse = await streamFromAI(feedbackMessages, () => S.isStopped);
  if (!improveResponse || S.isStopped) {
    const improveTiming = Date.now() - improveStartTime;

    try {
      trackPhase("improvement", false, improveTiming);
    } catch (e) {
      // Non-critical
    }

    submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
    onAcknowledge();
    return;
  }

  S.conversationHistory.push({ role: "assistant", content: improveResponse });
  const improveOps = extractOperations(improveResponse);

  if (!improveOps || improveOps.operations.length === 0) {
    appendMessage("agent", `ℹ ${improveResponse.substring(0, 300)}`);

    const improveTiming = Date.now() - improveStartTime;

    try {
      trackPhase("improvement", false, improveTiming);
    } catch (e) {
      // Non-critical
    }

    submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
    onAcknowledge();
    return;
  }

  // Execute improvement operations
  const improveStepStart = Date.now();
  const improveResults = [];
  let improveHadError = false;

  for (const op of improveOps.operations) {
    try {
      const runFn = S.currentHost === "powerpoint" ? PowerPoint.run : Excel.run;
      const result = await runFn(async (context) => {
        return executeOperationWithTracking(context, op);
      });
      improveResults.push(result);
    } catch (error) {
      improveResults.push(`✗ ${op.operation} failed: ${error.message}`);
      improveHadError = true;
    }
  }

  const improveTiming = Date.now() - improveStepStart;

  if (improveHadError) {
    appendMessage("agent", `❌ Some improvements failed: ${improveResults.join("\n")}`);

    // Track improvement phase failure
    try {
      trackPhase("improvement", false, improveTiming);
    } catch (e) {
      // Non-critical
    }

    submitFeedback("improve", comment, { ...taskStats, improvementApplied: false });
    onAcknowledge();
    return;
  }

  // Track improvement phase success
  try {
    trackPhase("improvement", true, improveTiming);
  } catch (e) {
    // Non-critical
  }

  const actualNames = getHostExtractNamesFn()(improveResults, improveOps.operations);
  const improvementVerification = await getHostVerifyFn()(improveOps.operations, actualNames);
  appendMessage(
    "agent",
    `✅ Improvements applied.${improvementVerification ? "\n" + improvementVerification : ""}`
  );

  trackStep({
    stepNumber: S.feedbackTotalSteps + 1,
    userPrompt: S.originalQuery,
    plan: `Improvement: ${comment}`,
    operations: improveOps.operations,
    results: improveResults,
    errors: [],
    success: true,
    timingMs: improveTiming,
    aiResponse: improveResponse,
    verification: improvementVerification || "",
    conversationHistory: [...S.conversationHistory],
    sheetContext,
    systemPrompt,
    aiModel: getConfig().model,
    aiEndpoint: getConfig().endpoint,
    phase: "improvement",
  });
  S.feedbackTotalSteps++;
  S.feedbackTotalTiming += improveTiming;

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
        onAcknowledge();
      },
      onBad: () => {
        submitFeedback("bad", comment, { ...taskStats, improvementApplied: true });
        onAcknowledge();
      },
      onImprove: (newComment) => {
        handleImprove(newComment, { ...taskStats, improvementApplied: true }, onAcknowledge);
      },
    });
  }

  scrollChatToBottom();
}

// ─── Main Send Handler ──────────────────────────────────────────────────────

async function handleSend(onAcknowledge, onImprove) {
  await loadHostModules();
  const chatInput = document.getElementById("chat-input");
  const message = chatInput.value.trim();
  if (!message && !S.pendingImageBase64) return;
  if (S.isExecuting) return;

  chatInput.value = "";
  chatInput.style.height = "auto";

  if (S.pendingImageBase64) {
    // Append user message with image
    const messagesContainer = document.getElementById("chat-messages");
    if (messagesContainer) {
      const messageEl = document.createElement("div");
      messageEl.className = "agent-message user";

      const bubbleEl = document.createElement("div");
      bubbleEl.className = "agent-bubble";

      const textEl = document.createElement("div");
      textEl.className = "user-message-text";
      if (message) {
        textEl.textContent = message;
        bubbleEl.appendChild(textEl);
      }

      const imgEl = document.createElement("img");
      imgEl.src = S.pendingImageBase64;
      imgEl.className = "user-message-image";
      imgEl.alt = "Attached image";
      bubbleEl.appendChild(imgEl);

      messageEl.appendChild(bubbleEl);
      messagesContainer.appendChild(messageEl);
      scrollChatToBottom();
    }

    S.conversationHistory.push({
      role: "user",
      content: [
        { type: "text", text: message },
        { type: "image_url", image_url: { url: S.pendingImageBase64 } },
      ],
    });
    console.log("[agentChat] Sending with image to AI");
    S.pendingImageBase64 = null;
  } else {
    appendMessage("user", message);
    S.conversationHistory.push({ role: "user", content: message });
  }

  S.originalQuery = message;
  S.feedbackStepData = null;
  S.feedbackTotalSteps = 0;
  S.feedbackTotalTiming = 0;
  const taskStartTime = Date.now();

  // Capture before snapshot for telemetry
  try {
    const beforeContext = await getHostMetadataFn()();
    if (beforeContext) {
      trackSnapshot("before", beforeContext);
    }
  } catch {
    // Non-critical
  }

  S.isExecuting = true;
  setExecutingState(true);
  S.isStopped = false;
  const sendButton = document.getElementById("chat-send");
  const stopButton = document.getElementById("chat-stop");
  if (sendButton) sendButton.disabled = true;

  try {
    if (stopButton) stopButton.style.display = "inline-block";

    // Phase 1: Planning
    console.log("[plan] === Starting planning phase ===");
    let sheetContext = await getHostContextFn()();
    console.log("[plan] Context:", sheetContext);
    let systemPrompt = buildSystemPrompt(S.currentHost, S.currentMode, sheetContext);

    const apiMessages = [
      { role: "system", content: systemPrompt },
      ...(await summarizeConversationHistory(S.conversationHistory, streamFromAI)),
    ];
    console.log(
      "[plan] Sending to AI — system prompt length:",
      systemPrompt.length,
      "msg count:",
      apiMessages.length
    );

    const stream = appendStreamingMessage();
    let planResponse = "";

    try {
      // Track planning phase start
      try {
        trackPhase("planning", true, 0);
      } catch (e) {
        // Non-critical
      }

      planResponse = await streamFromAI(
        apiMessages,
        () => S.isStopped,
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
      const planTiming = Date.now() - taskStartTime;
      stream.element.remove();
      appendMessage("agent", `Error: ${error.message || String(error)}`);

      // Track AI API failure with full context
      trackStep({
        stepNumber: 0,
        userPrompt: S.originalQuery,
        plan: "",
        operations: [],
        results: [],
        errors: [`AI API error: ${error.message || String(error)}`],
        success: false,
        timingMs: planTiming,
        aiResponse: "",
        verification: "",
        conversationHistory: apiMessages
          ? apiMessages.map((m) => ({
              role: m.role,
              content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
            }))
          : [],
        sheetContext,
        systemPrompt,
        aiModel: getConfig().model,
        aiEndpoint: getConfig().endpoint,
        phase: "planning",
      });

      throw error;
    }

    console.log("[plan] AI response received (length:", planResponse.length, ")");
    S.conversationHistory.push({ role: "assistant", content: planResponse });

    // Show plan
    appendPlanMessage(planResponse);

    // Extract operations
    let finalOps = extractOperations(planResponse);
    console.log("[plan] Extracted ops:", finalOps);

    if (!finalOps || finalOps.operations.length === 0) {
      // No operations — check if AI is claiming completion
      if (isCompletionClaim(planResponse)) {
        console.log("[plan] AI claims complete without operations");
        appendMessage("agent", planResponse);
        S.conversationHistory = [];
        S.stepStack = [];
        S.originalQuery = "";
        return;
      }

      // No operations — ask for clarification
      appendMessage("agent", "⚠ No operations found. Asking to continue...");
      const retryFeedback =
        "Your last response did not contain structured operations. Please provide the next step as structured operations, or say 'All steps complete.' if done.";
      S.conversationHistory.push({ role: "user", content: retryFeedback });

      const retryResponse = await streamFromAI(
        [{ role: "system", content: systemPrompt }, ...S.conversationHistory],
        () => S.isStopped
      );
      if (!retryResponse || S.isStopped) return;

      S.conversationHistory.push({ role: "assistant", content: retryResponse });
      finalOps = extractOperations(retryResponse);
      if (!finalOps || finalOps.operations.length === 0) {
        appendMessage("agent", "⚠ Agent could not generate operations. Stopping.");
        return;
      }
    }

    // Validate operations
    const hostModulesState = getHostModules();
    const registry =
      S.currentHost === "powerpoint" && hostModulesState?.ops
        ? hostModulesState.ops.PPT_OPERATION_REGISTRY
        : EXCEL_OPERATION_REGISTRY;
    const { valid: validOps, errors: validationErrors } = validateOperations(
      finalOps.operations,
      registry,
      S.currentMode
    );
    console.log(
      "[plan] Validation — valid:",
      validOps.length,
      "errors:",
      validationErrors.length,
      validationErrors
    );

    if (validationErrors.length > 0) {
      // Try to fix validation errors
      console.warn("[plan] Validation failed, retrying with fix prompt:", validationErrors);

      const fixMessages = [
        { role: "system", content: systemPrompt },
        { role: "assistant", content: planResponse },
        {
          role: "user",
          content:
            "Your operations failed validation. Please fix the following issues and respond with a corrected ```json block:\n" +
            validationErrors.join("\n") +
            "\n\nEnsure each operation has ALL required fields with correct types. Respond with ONLY the corrected JSON block.",
        },
      ];

      try {
        const fixResponse = await streamFromAI(
          fixMessages,
          () => S.isStopped,
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
          const retryOps = extractOperations(fixResponse);
          if (retryOps && retryOps.operations.length > 0) {
            const { valid: retryValidOps, errors: retryErrors } = validateOperations(
              retryOps.operations,
              registry,
              S.currentMode
            );
            if (retryErrors.length === 0) {
              console.log(
                "[plan] Validation retry succeeded —",
                retryValidOps.length,
                "operations"
              );
              finalOps = retryOps;
              planResponse = fixResponse;
            } else {
              console.warn("[plan] Validation retry also failed:", retryErrors);
              stream.element.remove();
              appendMessage("agent", `⚠ Retry failed: ${retryErrors.join("; ")}`);
              return;
            }
          } else {
            console.warn("[plan] Retry produced no operations");
            stream.element.remove();
            appendMessage("agent", "⚠ Retry produced no valid operations.");
            return;
          }
        }
      } catch (retryError) {
        console.error("[plan] Validation retry error:", retryError);
        stream.element.remove();
        appendMessage("agent", `⚠ Retry failed: ${retryError.message || String(retryError)}`);
        return;
      }
    }

    // Phase 2: Execute operations loop
    await executeOperationsLoop(finalOps, sheetContext, systemPrompt, onAcknowledge, onImprove);
  } catch (error) {
    const handleTiming = Date.now() - taskStartTime;
    console.error("[agentChat] handleSend error:", error);
    stream?.element?.remove();
    appendMessage("agent", `Error: ${error.message || String(error)}`);

    // Track unexpected error with full context
    trackStep({
      stepNumber: 0,
      userPrompt: S.originalQuery,
      plan: "",
      operations: [],
      results: [],
      errors: [`Unexpected error: ${error.message || String(error)}`],
      success: false,
      timingMs: handleTiming,
      aiResponse: "",
      verification: "",
      conversationHistory: apiMessages
        ? apiMessages.map((m) => ({
            role: m.role,
            content: typeof m.content === "string" ? m.content.substring(0, 500) : m.content,
          }))
        : [],
      sheetContext,
      systemPrompt,
      aiModel: getConfig().model,
      aiEndpoint: getConfig().endpoint,
      phase: "planning",
    });
  } finally {
    S.isExecuting = false;
    setExecutingState(false);
    S.isStopped = false;
    if (sendButton) sendButton.disabled = false;
    if (stopButton) stopButton.style.display = "none";
    chatInput.focus();
  }
}

// Export for use by init and agentChat
export {
  handleSend,
  handleRetry,
  handleRevert,
  handleImprove,
  continueExecutionFromAI,
  executeOperationsLoop,
};
