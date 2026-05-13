/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, PowerPoint, Office, localStorage */

// Centralized host detection — runs once at startup
let currentHost = null;

Office.onReady((info) => {
  currentHost = info.host;
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  console.log(`[taskpane] Running in host: ${currentHost}`);

  // Store host in localStorage for agentChat to use
  const pptHost = currentHost === Office.HostType.PowerPoint ? "powerpoint" : "excel";
  localStorage.setItem("agentHost", pptHost);

  const runBtn = document.getElementById("run");
  if (runBtn) {
    runBtn.onclick = run;
  }

  // Initialize the AI Agent chat
  import("./agentChat.js")
    .then((module) => {
      console.log("[taskpane] agentChat module loaded");
      module.initAgentChat();
      console.log("[taskpane] initAgentChat called");
    })
    .catch((err) => {
      console.error("[taskpane] Failed to load agentChat:", err);
    });
});

/**
 * Get the current Office host (Excel or PowerPoint).
 * @returns {"excel"|"powerpoint"}
 */
export function getCurrentHost() {
  return currentHost === Office.HostType.PowerPoint ? "powerpoint" : "excel";
}

export async function run() {
  try {
    if (currentHost === Office.HostType.PowerPoint) {
      await PowerPoint.run(async (context) => {
        const selection = context.presentation.getSelectedShapes();
        selection.load("name");
        await context.sync();
        console.log(`Selected: ${selection.items.length} shape(s)`);
      });
    } else {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("address");
        range.format.fill.color = "yellow";
        await context.sync();
        console.log(`The range address was ${range.address}.`);
      });
    }
  } catch (error) {
    console.error(error);
  }
}
