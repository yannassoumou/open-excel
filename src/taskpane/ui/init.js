/* global console, document, localStorage */

/* eslint-disable no-undef */

/* Chat UI initialization — DOM setup, event wiring, image preview */

import { S } from "../store.js";
import { setConfig as setAiConfig, getConfig, fetchModels, DEFAULTS } from "../agent/ai.js";
import { appendMessage, showWelcome, scrollChatToBottom, escapeHtml } from "./chat.js";
import { parseTSV, showPastePreview } from "./paste.js";
import { switchMode as switchModeUI } from "./config.js";
import { init as initTelemetry, flush as flushTelemetry } from "../telemetry.js";
import { handleSend } from "../agent/executor.js";

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
      S.pendingImageBase64 = null;
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

// ─── Initialization ──────────────────────────────────────────────────────────

export async function initAgentChat(onAcknowledge, onImprove) {
  try {
    console.log("[agentChat] initAgentChat called");

    // Ensure host is set
    if (!localStorage.getItem("agentHost")) {
      localStorage.setItem("agentHost", S.currentHost);
    }
    S.currentHost = localStorage.getItem("agentHost") || "excel";

    const appBody = document.getElementById("app-body");
    console.log("[agentChat] appBody:", appBody, "host:", S.currentHost);
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
          <span id="mode-badge" style="font-size:12px;font-weight:600;">${S.currentMode === "interactive" ? "⚡ Edit Mode" : "📖 Read Only"}</span>
          <div class="agent-mode-toggle">
            <button id="mode-interactive" class="agent-mode-btn ${S.currentMode === "interactive" ? "mode-active" : ""}">⚡ Edit Mode</button>
            <button id="mode-explain" class="agent-mode-btn ${S.currentMode === "explain" ? "mode-active" : ""}">📖 Read Only</button>
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
          <p>${S.currentHost === "powerpoint" ? "Describe what you want to do in PowerPoint. The agent will generate and execute structured operations automatically." : "Tell Kuro what you need — it figures out the rest."}</p>
          <p style="margin-top: 12px; font-size: 12px; color: #999;">Examples:<br/>
            ${
              S.currentHost === "powerpoint"
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
        <textarea id="chat-input" placeholder="Ask Kuro anything about your ${S.currentHost === "powerpoint" ? "slides" : "spreadsheet"}..." rows="1"></textarea>
        <input type="file" id="chat-image-input" accept="image/*" style="display:none;" />
        <div class="input-btn-group">
          <button id="chat-image-upload" class="input-btn input-btn-image" title="Attach image" style="display:none;">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21.44 11.05l-9.19 9.19a6 6 0 01-8.49-8.49l9.19-9.19a4 4 0 015.66 5.66l-9.2 9.19a2 2 0 01-2.83-2.83l8.49-8.48"/></svg>
          </button>
          <button id="chat-send" class="input-btn input-btn-send" disabled title="Send message">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="22" y1="2" x2="11" y2="13"/><polygon points="22 2 15 22 11 13 2 9 22 2"/></svg>
          </button>
          <button id="chat-stop" class="input-btn input-btn-stop" style="display:none;" title="Stop execution">
            <svg viewBox="0 0 24 24" fill="currentColor"><rect x="4" y="4" width="16" height="16" rx="2"/></svg>
          </button>
          <button id="chat-clear" class="input-btn input-btn-clear" title="Clear chat">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>
          </button>
        </div>
      </div>
    `;

    document.body.appendChild(chatContainer);

    // Restyle app-body with new welcome content
    const host = S.currentHost === "powerpoint" ? "PowerPoint" : "Excel";
    const hostLower = S.currentHost === "powerpoint" ? "PowerPoint" : "Excel";
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

    // Preset models shown as fallback before API discovery
    const PRESET_MODELS = [
      "nvidia/nemotron-3-super-120b-a12b:free",
      "openai/gpt-4",
      "openai/gpt-4o",
      "google/gemini-2.5-pro-preview-05-06",
      "meta-llama/llama-3.3-70b-instruct",
      "mistralai/mistral-large-2-instruct",
      "deepseek/deepseek-chat",
      "nousresearch/hermes-3-llama-3.1-70b",
      "local-model",
    ];

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
          // Build datalist: presets + discovered models (no duplicates)
          const datalist = document.getElementById("model-suggestions");
          if (datalist) {
            datalist.innerHTML = "";

            // Add presets first (only if not already in discovered models)
            PRESET_MODELS.forEach((v) => {
              const opt = document.createElement("option");
              opt.value = v;
              datalist.appendChild(opt);
            });

            // Add discovered models (avoid duplicates with presets)
            models.forEach((m) => {
              if (!PRESET_MODELS.includes(m)) {
                const opt = document.createElement("option");
                opt.value = m;
                datalist.appendChild(opt);
              }
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
        // Clear model when endpoint changes
        if (modelInput) modelInput.value = "";
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
      interactiveModeBtn.addEventListener("click", () => {
        const prev = S.currentMode;
        const next = switchModeUI("interactive", appendMessage) || prev;
        S.currentMode = next;
      });
    }
    if (explainModeBtn) {
      explainModeBtn.addEventListener("click", () => {
        const prev = S.currentMode;
        const next = switchModeUI("explain", appendMessage) || prev;
        S.currentMode = next;
      });
    }

    if (chatInput) {
      chatInput.addEventListener("input", () => {
        chatInput.style.height = "auto";
        chatInput.style.height = Math.min(chatInput.scrollHeight, 120) + "px";
        sendButton.disabled = chatInput.value.trim() === "" && !S.isExecuting;
      });
    }

    if (chatInput) {
      chatInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          handleSend(onAcknowledge, onImprove);
        } else if (e.key === "Tab") {
          e.preventDefault();
          const nextMode = S.currentMode === "interactive" ? "explain" : "interactive";
          const prev = S.currentMode;
          const next = switchModeUI(nextMode, appendMessage) || prev;
          S.currentMode = next;
          chatInput.placeholder =
            S.currentMode === "interactive"
              ? `Ready to modify – describe what to do in ${S.currentHost === "powerpoint" ? "PowerPoint" : "Excel"}...`
              : "Read-only – ask about your content (Tab to switch)...";
          clearTimeout(chatInput._placeholderTimeout);
          chatInput._placeholderTimeout = setTimeout(() => {
            chatInput.placeholder = `Ask Kuro anything about your ${S.currentHost === "powerpoint" ? "slides" : "spreadsheet"}...`;
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
              S.conversationHistory.push({
                role: "assistant",
                content: `Wrote ${rows}×${cols} data to active ${S.currentHost === "powerpoint" ? "slide" : "worksheet"}.`,
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
      sendButton.addEventListener("click", () => {
        handleSend(onAcknowledge, onImprove);
      });
    }

    if (stopButton) {
      stopButton.addEventListener("click", () => {
        S.isExecuting = false;
        S.isStopped = true;
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
          S.pendingImageBase64 = event.target.result;
          showImagePreview(S.pendingImageBase64);
          imageUploadBtn.style.display = "inline-block";
          console.log("[agentChat] Image attached:", file.name);
        };
        reader.readAsDataURL(file);
        imageInput.value = "";
      });
    }

    if (imageUploadBtn) {
      imageUploadBtn.addEventListener("click", () => {
        if (S.pendingImageBase64) {
          S.pendingImageBase64 = null;
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
        if (S.isExecuting) {
          S.isExecuting = false;
          S.isStopped = true;
          document.getElementById("chat-stop").style.display = "none";
          sendButton.disabled = false;
        }
        // Import and call clearChat from agentChat
        import("../agentChat.js").then((mod) => mod.clearChat());
      });
    }

    // ─── Config Toggle & Auto-Deflate ────────────────────────────────────────

    const configBar = document.getElementById("agent-config");
    const configToggle = document.getElementById("config-toggle");
    let configExpanded = localStorage.getItem("agentConfigExpanded");
    if (configExpanded === null) configExpanded = true; // Show on first visit
    configExpanded = configExpanded === "true";
    let autoDeflateTimer = null;
    const AUTO_DEFDELAY = 60000; // 60 seconds

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
      if (configExpanded && S.isExecuting) return; // Don't auto-deflate during execution
      if (configExpanded) {
        autoDeflateTimer = setTimeout(() => {
          if (configExpanded && !S.isExecuting) {
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
      input.addEventListener("blur", () => {
        // Only auto-deflate when focus leaves the entire config area
        setTimeout(() => {
          if (configExpanded && !configBar.contains(document.activeElement)) {
            resetAutoDeflate();
          }
        }, 100);
      });
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
    try {
      await initTelemetry(aiConf.model, DEFAULTS.telemetryEndpoint);
    } catch (err) {
      console.error("[agentChat] Telemetry init failed:", err);
    }

    console.log("[agentChat] initAgentChat complete, host:", S.currentHost);
  } catch (err) {
    console.error("[agentChat] Error during initialization:", err);
  }
}
