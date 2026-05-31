/* Host module loader and context accessor functions */

import { S, getHost } from "../store.js";

let hostModules = null;

export function getHostModules() {
  return hostModules;
}

export async function loadHostModules() {
  if (hostModules) return hostModules;

  if (S.currentHost === "powerpoint") {
    const [pptContext, pptSnapshot, pptOps] = await Promise.all([
      import("../ppt/context.js"),
      import("../ppt/snapshot.js"),
      import("../ppt/operations.js"),
    ]);
    hostModules = {
      context: pptContext,
      snapshot: pptSnapshot,
      ops: pptOps,
    };
  } else {
    const [excelContext, excelSnapshot] = await Promise.all([
      import("../excel/context.js"),
      import("../excel/snapshot.js"),
    ]);
    hostModules = {
      context: excelContext,
      snapshot: excelSnapshot,
      ops: null,
    };
  }

  return hostModules;
}

export function getHostContextFn() {
  return getHost() === "powerpoint"
    ? hostModules.context.getSlideContext
    : hostModules.context.getSheetContext;
}

export function getHostFullContextFn() {
  return getHost() === "powerpoint"
    ? hostModules.context.readFullPresentationContext
    : hostModules.context.readFullWorkbookContext;
}

export function getHostMetadataFn() {
  return getHost() === "powerpoint"
    ? hostModules.context.readPresentationMetadata
    : hostModules.context.readWorkbookMetadata;
}

export function getHostCaptureSnapshotFn() {
  return getHost() === "powerpoint"
    ? hostModules.snapshot.captureSnapshot
    : hostModules.snapshot.captureSnapshot;
}

export function getHostVerifyFn() {
  return getHost() === "powerpoint"
    ? hostModules.snapshot.verifyPptOperations
    : hostModules.snapshot.verifyOperations;
}

export function getHostExtractNamesFn() {
  return getHost() === "powerpoint"
    ? hostModules.snapshot.extractSlideNamesFromResults
    : hostModules.snapshot.extractSheetNamesFromResults;
}
