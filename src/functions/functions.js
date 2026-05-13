/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console localStorage, fetch */

// ─── Demo functions (always available) ───────────────────────────────────────

/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
function logMessage(message) {
  console.log(message);
  return message;
}

// ─── KuroAgent AI Custom Functions ───────────────────────────────────────────

const DEFAULT_ENDPOINT = "https://openrouter.ai/api/v1/chat/completions";
const DEFAULT_MODEL = "gpt-4";
const MAX_TOKENS = 4096;
const TEMPERATURE = 0.3;
const REQUEST_TIMEOUT = 300000; // 5 minutes

/**
 * Read KuroAgent config from localStorage (same origin as task pane).
 * @returns {{endpoint: string, model: string, apiKey: string}}
 */
function readConfig() {
  const endpoint = localStorage.getItem("agentEndpoint") || DEFAULT_ENDPOINT;
  const model = localStorage.getItem("agentModel") || DEFAULT_MODEL;
  const apiKey = localStorage.getItem("agentApiKey") || "";
  return { endpoint, model, apiKey };
}

/**
 * KuroAgent — Ask any LLM directly from Excel.
 * 
 * Uses the same config as the KuroAgent task pane (endpoint, API key, model).
 * Configure once in the task pane, then use =KUROAGENT() anywhere.
 * 
 * @customfunction KUROAGENT
 * @param {string} prompt Your question or instruction for the AI
 * @param {string} [context] Optional additional context (e.g., "fr→en" for translation)
 * @returns {string} AI response text
 * @example =KUROAGENT("Résume les tendances de vente de ce tableau")
 * @example =KUROAGENT("Traduis ce texte en anglais", "fr→en")
 * @example =KUROAGENT("Formule Excel pour calculer la croissance mensuelle")
 */
function kuroAgent(prompt, context) {
  const config = readConfig();
  
  // Validate
  if (!prompt || typeof prompt !== "string" || prompt.trim() === "") {
    return "⚠️ Veuillez fournir une question ou instruction.";
  }
  
  if (!config.apiKey || config.apiKey.includes("PLACEHOLDER") || config.apiKey === "") {
    return "⚠️ Clé API non configurée. Ouvrez le panneau KuroAgent → Settings → entrez votre clé API.";
  }
  
  // Build prompt with optional context
  const fullPrompt = context
    ? `${prompt}\n\nContext: ${context}`
    : prompt;
  
  // Build messages
  const messages = [
    {
      role: "system",
      content: "You are a helpful assistant integrated into Excel. Provide clear, concise answers. When asked about Excel formulas or operations, provide specific, usable responses."
    },
    {
      role: "user",
      content: fullPrompt
    }
  ];
  
  // Return a Promise — Excel custom functions support async
  return fetch(config.endpoint, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${config.apiKey}`
    },
    body: JSON.stringify({
      model: config.model,
      messages: messages,
      max_tokens: MAX_TOKENS,
      temperature: TEMPERATURE
    }),
    signal: AbortSignal.timeout(REQUEST_TIMEOUT)
  })
  .then(response => {
    if (!response.ok) {
      return response.text().then(errText => {
        throw new Error(`API error ${response.status}: ${errText.substring(0, 200)}`);
      });
    }
    return response.json();
  })
  .then(data => {
    const content = data.choices?.[0]?.message?.content;
    if (content) {
      // Truncate very long responses to fit in Excel cells
      const maxLen = 8192;
      return content.length > maxLen ? content.substring(0, maxLen) + "\n\n...(tronqué)" : content;
    }
    return "ℹ️ Aucune réponse de l'IA. Vérifiez votre clé API et le modèle.";
  })
  .catch(error => {
    if (error.name === "TimeoutError") {
      return "⏱️ Délai d'attente dépassé (5 min). Le modèle met trop de temps à répondre.";
    }
    if (error.message && error.message.includes("API error")) {
      return `❌ Erreur API: ${error.message.substring(0, 300)}`;
    }
    return `❌ Erreur: ${error.message || String(error)}`;
  });
}

/**
 * KuroAgent Stream — same as KUROAGENT but returns a loading status initially.
 * For long-running queries, shows progress.
 * 
 * @customfunction KUROAGENT_STREAM
 * @param {string} prompt Your question or instruction
 * @param {string} [context] Optional context
 * @returns {string} AI response (may be partial for very long responses)
 */
function kuroAgentStream(prompt, context) {
  // For now, same as kuroAgent — streaming in custom functions is complex
  // Future: could use CustomFunctions.StreamingInvocation for real streaming
  return kuroAgent(prompt, context);
}
