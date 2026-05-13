/* global console, sessionStorage, AbortController, fetch, TextDecoder, clearTimeout, setTimeout, setInterval, clearInterval */

/* AI communication — streaming chat completions */

const DEFAULT_ENDPOINT = "https://openrouter.ai/api/v1/chat/completions";
const DEFAULT_MODEL = "gpt-4";
const TELEMETRY_ENDPOINT_DEFAULT = "https://telemetry-ivory.vercel.app/api/v1";
const APP_VERSION = "1.0.2";

let configuredEndpoint = sessionStorage.getItem("agentEndpoint") || DEFAULT_ENDPOINT;
let configuredModel = sessionStorage.getItem("agentModel") || DEFAULT_MODEL;
let configuredApiKey =
  sessionStorage.getItem("agentApiKey") || "sk-or-v1-PLACEHOLDER_ENTREZ_VOTRE_CLE";

export const DEFAULTS = {
  endpoint: DEFAULT_ENDPOINT,
  model: DEFAULT_MODEL,
  telemetryEndpoint: TELEMETRY_ENDPOINT_DEFAULT,
  version: APP_VERSION,
};

export function getEndpoint() {
  return configuredEndpoint;
}

export function getApiKey() {
  return configuredApiKey;
}

export function getConfig() {
  return {
    endpoint: configuredEndpoint,
    model: configuredModel,
    apiKey: configuredApiKey,
    version: APP_VERSION,
  };
}

/**
 * Update configuration (called from UI on config input changes).
 * Also persists to sessionStorage and returns the updated config.
 */
export function setConfig(updates) {
  if (updates.endpoint !== undefined) {
    configuredEndpoint = updates.endpoint || DEFAULT_ENDPOINT;
    sessionStorage.setItem("agentEndpoint", configuredEndpoint);
  }
  if (updates.model !== undefined) {
    configuredModel = updates.model || DEFAULT_MODEL;
    sessionStorage.setItem("agentModel", configuredModel);
  }
  if (updates.apiKey !== undefined) {
    configuredApiKey = updates.apiKey;
    sessionStorage.setItem("agentApiKey", configuredApiKey);
  }
  return getConfig();
}

/**
 * Stream response from AI endpoint using SSE.
 * @param {Array} messages - Chat messages
 * @param {function} onStop - Sync check — return truthy to abort
 * @param {number} maxTokens - Maximum tokens (default 8192)
 * @param {function} onChunk - Optional callback for each chunk (text, isComplete)
 * @returns {Promise<string|null>} Full AI response text
 */
let _activeController = null;

export function abortActiveRequest() {
  if (_activeController) {
    _activeController.abort();
    _activeController = null;
  }
}

export async function streamFromAI(messages, onStop, maxTokens = 16384, onChunk = null) {
  const headers = { "Content-Type": "application/json" };
  if (configuredApiKey) {
    headers["Authorization"] = `Bearer ${configuredApiKey}`;
  }

  const controller = new AbortController();
  _activeController = controller;
  const timeout = setTimeout(() => controller.abort(), 300000); // 5 min timeout

  try {
    // Notify UI that request is starting
    if (onChunk) onChunk("", false);

    const response = await fetch(configuredEndpoint, {
      method: "POST",
      headers,
      body: JSON.stringify({
        model: configuredModel,
        messages,
        max_tokens: maxTokens,
        temperature: 0.3,
      }),
      signal: controller.signal,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API error ${response.status}: ${errorText}`);
    }

    const data = await response.json();
    const result = data.choices?.[0]?.message?.content?.trim() || null;

    // Check if user aborted while we were waiting
    if (onStop && onStop()) {
      console.log("[agent:ai] Request interrupted by user");
      return "";
    }

    if (onChunk && result) onChunk(result, true);
    return result;
  } finally {
    clearTimeout(timeout);
    _activeController = null;
  }
}

/**
 * Discover available models from an OpenAI-compatible API endpoint.
 * Tries <baseUrl>/v1/models with the provided API key.
 * Handles multiple response formats (OpenAI, Ollama, NousResearch).
 * @param {string} endpoint - The chat completions endpoint URL
 * @param {string} apiKey - Bearer token for authentication
 * @returns {Promise<string[]>} Array of model ID strings
 */
export async function fetchModels(endpoint, apiKey) {
  let baseUrl = endpoint;

  // Strip common suffixes to get base URL
  const suffixes = ["/v1/chat/completions", "/chat/completions", "/v1", ""];
  for (const suffix of suffixes) {
    if (baseUrl.endsWith(suffix) && suffix.length > 0) {
      baseUrl = baseUrl.slice(0, -suffix.length);
      break;
    }
  }

  // Remove trailing slash
  baseUrl = baseUrl.replace(/\/+$/, "");

  const modelsUrl = `${baseUrl}/v1/models`;

  const headers = { "Content-Type": "application/json" };
  if (apiKey) {
    headers["Authorization"] = `Bearer ${apiKey}`;
  }

  const response = await fetch(modelsUrl, { headers });

  if (!response.ok) {
    throw new Error(`Models endpoint returned ${response.status}`);
  }

  const data = await response.json();

  // OpenAI-compatible format: { data: [{ id: "..." }] }
  if (data.data && Array.isArray(data.data)) {
    return data.data.map((m) => m.id).filter(Boolean);
  }

  // Ollama format: { models: [{ name: "..." }] }
  if (data.models && Array.isArray(data.models)) {
    return data.models.map((m) => m.name || m.id).filter(Boolean);
  }

  // NousResearch / OpenRouter format: { data: [{ id: "..." }] }
  if (data.data && typeof data.data === "object") {
    return Object.values(data.data)
      .map((m) => (typeof m === "string" ? m : m.id))
      .filter(Boolean);
  }

  return [];
}
