import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  applyExcelAdvancedAction,
  disposeOfficeWorker,
  DotNetOfficeActionBridge,
} from "./officeWorkerSmokeHelpers";
import type { OfficeActionInput, OfficeActionResult } from "../electron/agent/tools/officeCore/types";

async function main(): Promise<void> {
  const root = await mkdtemp(path.join(os.tmpdir(), "wengge-word-bookmark-reliability-"));
  const sourcePath = path.join(root, "source.xlsx");
  const wordPath = path.join(root, "linked.docx");
  const linkId = "bookmark-reliability";
  const bridge = new DotNetOfficeActionBridge();
  try {
    const created = await applyExcelAdvancedAction({
      operation: "createWorkbook",
      filePath: sourcePath,
      params: { sheetNames: ["Sheet1"], values: [["部门", "金额"], ["华东", 10], ["华南", 20]] },
    });
    if (created.status !== "done") throw new Error(created.error || created.summary);
    await run(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToWord", filePath: sourcePath, outputPath: wordPath,
      target: "range:Sheet1!A1:B3", params: { linked: true, linkId, overwrite: true, sourceHost: "excel", wordHost: "word" },
    });
    await run(bridge, {
      app: "word", action: "insert", operation: "manageReferences", filePath: wordPath,
      params: { host: "word", command: "addBookmark", name: "ManualKeep" },
    });
    await run(bridge, {
      app: "word", action: "edit", operation: "applyTrackedChanges", filePath: wordPath,
      params: { host: "word", edits: [{ command: "replaceBookmark", name: "ManualKeep", text: "人工保留段落" }], keepTracking: false },
    });
    assertBookmark(await run(bridge, {
      app: "word", action: "inspect", operation: "inspectReferences", filePath: wordPath, params: { host: "word" },
    }));
    const write = await applyExcelAdvancedAction({
      operation: "writeRange", filePath: sourcePath, target: "range:Sheet1!B2:B3", params: { values: [[100], [200]] },
    });
    if (write.status !== "done") throw new Error(write.error || write.summary);
    const update = await run(bridge, {
      app: "excel", action: "insert", operation: "exportRangeToWord", filePath: sourcePath, outputPath: wordPath,
      target: "range:Sheet1!A1:B3", params: { linked: true, updateExisting: true, linkId, sourceHost: "excel", wordHost: "word" },
    });
    assertBookmark(await run(bridge, {
      app: "word", action: "inspect", operation: "inspectReferences", filePath: wordPath, params: { host: "word" },
    }));
    process.stdout.write(`${JSON.stringify({ ok: true, update: data(update) }, null, 2)}\n`);
  } finally {
    await disposeOfficeWorker();
    await rm(root, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

async function run(bridge: DotNetOfficeActionBridge, input: OfficeActionInput): Promise<OfficeActionResult> {
  process.stdout.write(`Testing ${input.app}/${input.operation}\n`);
  const result = await bridge.executeAction(input);
  if (result.status !== "done") throw new Error(`${input.operation}: ${result.error || result.summary}`);
  process.stdout.write(`Passed ${input.app}/${input.operation}\n`);
  return result;
}

function assertBookmark(result: OfficeActionResult): void {
  const bookmarks = Array.isArray(data(result).bookmarks) ? data(result).bookmarks as Array<Record<string, unknown>> : [];
  const bookmark = bookmarks.find((item) => item.name === "ManualKeep");
  if (!bookmark || !String(bookmark.text || "").includes("人工保留段落")) {
    throw new Error(`ManualKeep 书签未保留: ${JSON.stringify(bookmarks)}`);
  }
}

function data(result: OfficeActionResult): Record<string, unknown> {
  return result.data && typeof result.data === "object" && !Array.isArray(result.data)
    ? result.data as Record<string, unknown>
    : {};
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
