import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DotNetOfficeActionBridge as OfficeComActionBridge,
  applyExcelAdvancedAction,
  applyPresentationAdvancedAction,
  applyWordAdvancedAction,
  disposeOfficeWorker,
} from "./officeWorkerSmokeHelpers";
import type { OfficeActionInput, OfficeActionResult } from "../electron/agent/tools/officeCore/types";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-wps-routing-"));
  const excelPath = path.join(tempDir, "routing.xlsx");
  const wordPath = path.join(tempDir, "routing.docx");
  const presentationPath = path.join(tempDir, "routing.pptx");
  try {
    logStage("创建 Excel/Word/PowerPoint 测试文件");
    await createFixtures(excelPath, wordPath, presentationPath);
    logStage("测试文件创建完成，开始通过 WPS 打开现有文件");
    const bridge = new OfficeComActionBridge();
    const checks: Array<{ action: OfficeActionInput; expectedProgIds: string[] }> = [
      {
        action: { app: "excel", action: "inspect", operation: "inspectPrintSettings", filePath: excelPath, params: { host: "wps" } },
        expectedProgIds: ["ket.application"],
      },
      {
        action: { app: "word", action: "inspect", operation: "inspectDocumentFormatting", filePath: wordPath, params: { host: "wps" } },
        expectedProgIds: ["kwps.application", "wps.application"],
      },
      {
        action: { app: "presentation", action: "inspect", operation: "inspectPresentationTheme", filePath: presentationPath, params: { host: "wps" } },
        expectedProgIds: ["wpp.application", "kwpp.application"],
      },
    ];
    const routes = [];
    for (const check of checks) {
      logStage(`检查 ${check.action.app} WPS 路由`);
      const result = await bridge.executeAction(check.action);
      assertDone(result, `${check.action.app} WPS 路由检查`);
      const progId = String(asRecord(result.data).progId || "").toLowerCase();
      if (!check.expectedProgIds.includes(progId)) throw new Error(`${check.action.app} 未使用预期 WPS COM: ${progId || "unknown"}`);
      routes.push({ app: check.action.app, progId });
      logStage(`${check.action.app} WPS 路由通过: ${progId}`);
    }
    process.stdout.write(`${JSON.stringify({ ok: true, createdBeforeWpsOpen: true, routes }, null, 2)}\n`);
  } finally {
    await disposeOfficeWorker();
    await rm(tempDir, { recursive: true, force: true, maxRetries: 20, retryDelay: 300 });
  }
}

async function createFixtures(excelPath: string, wordPath: string, presentationPath: string): Promise<void> {
  const results = await Promise.all([
    applyExcelAdvancedAction({ operation: "createWorkbook", filePath: excelPath, params: { sheetNames: ["Sheet1"], values: [["Name", "Value"], ["WPS", 1]] } }),
    applyWordAdvancedAction({ operation: "createDocument", filePath: wordPath, params: { title: "WPS 路由测试", paragraphs: ["文件先创建，再由 WPS 打开。"] } }),
    applyPresentationAdvancedAction({ operation: "createPresentation", filePath: presentationPath, params: { title: "WPS 路由测试" } }),
  ]);
  for (const [index, result] of results.entries()) assertDone(result, `创建 WPS 路由夹具 ${index + 1}`);
}

function assertDone(result: OfficeActionResult, label: string): void {
  if (result.status !== "done") throw new Error(`${label}失败: ${result.error || result.summary}`);
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value) ? value as Record<string, unknown> : {};
}

function logStage(message: string): void {
  process.stdout.write(`[wps-routing] ${message}\n`);
}

void main().then(
  () => process.exit(0),
  (error) => {
    process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
    process.exit(1);
  },
);
