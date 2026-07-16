import { copyFile, mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DotNetExcelBridge } from "../electron/agent/officeWorker/dotNetExcelBridge";
import {
  applyExcelAdvancedAction,
  closeOfficeFixtures,
  disposeOfficeWorker,
  getExcelDisplayAlerts,
  listOfficeSmokeProcesses,
  openOfficeFixtures,
  runningOfficeSmokeProcesses,
  setExcelDisplayAlerts,
  setExcelStructureProtected,
} from "./officeWorkerSmokeHelpers";

type HostMode = "excel" | "wps" | "both";

type OwnedExcel = { excel: number[]; word: number[]; presentation: number[] };

const SHEETS = ["S1", "S2", "S3"] as const;
const STRUCTURE_PASSWORD = "wengge-smoke-m09";

async function main(): Promise<void> {
  process.env.WENGGE_OFFICE_SMOKE = "1";
  const hostMode = parseHostMode(process.env.WENGGE_EXCEL_DISPLAY_ALERTS_HOST);
  const summary: Record<string, unknown> = { hostMode, scenarios: [] as unknown[] };
  let tempDir: string | undefined;

  try {
    const processesBefore = await listOfficeSmokeProcesses();
    const microsoftBefore = processesBefore.microsoft;
    const wpsVisibleBefore = processesBefore.wpsVisible;
    const wpsAllBefore = processesBefore.wpsAll ?? wpsVisibleBefore;

    if ((hostMode === "excel" || hostMode === "both") && microsoftBefore.length > 0) {
      throw new Error(
        `DisplayAlerts 冒烟拒绝在已有 Microsoft Office 进程上运行（pids: ${microsoftBefore.join(", ")}）。请使用无用户 Office 窗口的隔离 Runner。`,
      );
    }
    if ((hostMode === "wps" || hostMode === "both") && wpsAllBefore.length > 0) {
      throw new Error(
        `WPS DisplayAlerts 冒烟拒绝附加已有 et/wps 表格进程（pids: ${wpsAllBefore.join(", ")}）。请使用无用户 WPS 表格的隔离 Runner。`,
      );
    }

    tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-excel-display-alerts-"));
    const excelPath = path.join(tempDir, "display-alerts-excel.xlsx");
    const wpsPath = path.join(tempDir, "display-alerts-wps.xlsx");
    await createFixture(excelPath);
    await assertNoHostStarted(microsoftBefore, wpsAllBefore, "createFixture(excel)");
    await copyFile(excelPath, wpsPath);
    await assertNoHostStarted(microsoftBefore, wpsAllBefore, "copyFixture(wps)");

    if (hostMode === "excel" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(await runExcelMatrix(excelPath, microsoftBefore));
    }
    if (hostMode === "wps" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(await runWpsMatrix(wpsPath, wpsAllBefore));
    }

    await assertProcessesStillRunning(microsoftBefore, "任务前 Microsoft Office 进程");
    await assertProcessesStillRunning(wpsVisibleBefore, "任务前可见 WPS 进程");
  } finally {
    await disposeOfficeWorker();
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
    }
  }

  process.stdout.write(`${JSON.stringify({ ok: true, ...summary }, null, 2)}\n`);
}

async function createFixture(filePath: string): Promise<void> {
  const result = await applyExcelAdvancedAction({
    operation: "createWorkbook",
    filePath,
    params: {
      sheetNames: [...SHEETS],
      values: [["marker", "m09"]],
    },
  });
  if (result.status !== "done") {
    throw new Error(`创建 DisplayAlerts 夹具失败: ${result.error || result.summary}`);
  }
}

async function runExcelMatrix(
  filePath: string,
  microsoftBefore: number[],
): Promise<Record<string, unknown>> {
  let owned: OwnedExcel | undefined;
  try {
    owned = await openOfficeFixtures({
      excelPaths: [filePath],
      wordPaths: [],
      presentationPaths: [],
    });
    const bridge = new DotNetExcelBridge();
    await attachToFixtureExcel(bridge, "Excel DisplayAlerts openFixtures");
    const cases = await runDisplayAlertsMatrix(bridge, "excel");
    const ownedIds = ownedIdsFrom(owned);
    await closeOfficeFixtures(owned);
    owned = undefined;
    await assertOwnedStopped(ownedIds);
    await assertNoUnexpectedMicrosoft(microsoftBefore);
    return { host: "excel", fixture: path.basename(filePath), ...cases };
  } finally {
    if (owned) {
      try {
        await closeOfficeFixtures(owned);
      } catch {
        /* best-effort */
      }
    }
  }
}

async function runWpsMatrix(
  filePath: string,
  wpsAllBefore: number[],
): Promise<Record<string, unknown>> {
  if (wpsAllBefore.length > 0) {
    throw new Error(
      `WPS DisplayAlerts 冒烟拒绝附加已有 et/wps 表格进程（pids: ${wpsAllBefore.join(", ")}）。`,
    );
  }

  const bridge = new DotNetExcelBridge();
  const selected = await bridge.selectHost("wps");
  const opened = await bridge.openWorkbook(filePath).catch((error: unknown) => ({
    success: false as const,
    error: error instanceof Error ? error.message : String(error),
  }));
  if (!opened.success) {
    throw new Error(
      `WPS 无法连接/打开工作簿: ${opened.error || `selectHost connected=${selected.connected}`}`,
    );
  }

  const afterOpen = await listOfficeSmokeProcesses();
  const wpsAllAfter = afterOpen.wpsAll ?? afterOpen.wpsVisible;
  const ownedWpsPids = wpsAllAfter.filter((id) => !wpsAllBefore.includes(id));
  if (ownedWpsPids.length === 0) {
    throw new Error("WPS openWorkbook 后未观察到本任务创建的 et/wps 表格进程");
  }

  try {
    const cases = await runDisplayAlertsMatrix(bridge, "wps");
    // Dispose so owned host exits; readonly PID check may briefly restart Worker; dispose again.
    await disposeOfficeWorker();
    await assertOwnedStopped(ownedWpsPids);
    await disposeOfficeWorker();
    return { host: "wps", fixture: path.basename(filePath), ownedWpsPids, ...cases };
  } finally {
    await assertProcessesStillRunning(wpsAllBefore, "任务前 WPS 表格进程");
  }
}

/**
 * Production path: DotNetExcelBridge.sheetOperation("delete") → excel.sheet.operation
 * → ExcelWorkbookService.SheetOperation delete (DisplayAlerts save/restore in finally).
 *
 * (a) original DisplayAlerts=false; successful delete restores false; then restore true
 * (b) original DisplayAlerts=true; structure-protect delete fails; still true; sheet remains
 * (c) original DisplayAlerts=true; last-visible delete fails; still true; sheet remains
 */
async function runDisplayAlertsMatrix(
  bridge: DotNetExcelBridge,
  host: "excel" | "wps",
): Promise<Record<string, unknown>> {
  const labels = {
    success: `${host}:alerts_false_success_restore`,
    structure: `${host}:structure_protect_fail_restore_true`,
    lastVisible: `${host}:last_visible_fail_restore_true`,
  };

  const success = await caseAlertsFalseSuccessRestore(bridge, labels.success);
  const structure = await caseStructureProtectFailRestore(bridge, labels.structure);
  await bridge.sheetOperation("delete", "S3");
  await assertSheetNames(bridge, ["S1"], `${labels.lastVisible}:pre`);
  const lastVisible = await caseLastVisibleFailRestore(bridge, labels.lastVisible);

  const finalAlerts = await getExcelDisplayAlerts();
  if (finalAlerts.displayAlerts !== true) {
    throw new Error(`${host}: 矩阵结束后 DisplayAlerts 应为 true: ${JSON.stringify(finalAlerts)}`);
  }

  return {
    alertsFalseSuccessRestore: success,
    structureProtectFailRestore: structure,
    lastVisibleFailRestore: lastVisible,
    displayAlertsFinalTrue: true,
  };
}

async function caseAlertsFalseSuccessRestore(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>> {
  await assertSheetNames(bridge, ["S1", "S2", "S3"], `${label}:before`);
  const set = await setExcelDisplayAlerts(false);
  if (set.displayAlerts !== false) {
    throw new Error(`${label}: 无法将 DisplayAlerts 设为 false: ${JSON.stringify(set)}`);
  }
  await bridge.sheetOperation("delete", "S2");
  const after = await getExcelDisplayAlerts();
  if (after.displayAlerts !== false) {
    throw new Error(`${label}: 成功删除后 DisplayAlerts 未恢复为 false: ${JSON.stringify(after)}`);
  }
  await assertSheetNames(bridge, ["S1", "S3"], `${label}:after`);

  // Task: after false-success assertion, immediately restore true for safe cleanup.
  const restored = await setExcelDisplayAlerts(true);
  if (restored.displayAlerts !== true) {
    throw new Error(
      `${label}: 断言后无法将 DisplayAlerts 恢复为 true: ${JSON.stringify(restored)}`,
    );
  }

  return {
    original: false,
    displayAlertsAfterDelete: after.displayAlerts,
    displayAlertsAfterRestoreTrue: restored.displayAlerts,
    deleted: "S2",
    sheets: ["S1", "S3"],
  };
}

async function caseStructureProtectFailRestore(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>> {
  await assertSheetNames(bridge, ["S1", "S3"], `${label}:before`);
  // Original value true — covers true-restore on failure path.
  const set = await setExcelDisplayAlerts(true);
  if (set.displayAlerts !== true) {
    throw new Error(`${label}: 无法将 DisplayAlerts 设为 true: ${JSON.stringify(set)}`);
  }

  const protectedState = await setExcelStructureProtected({
    protected: true,
    password: STRUCTURE_PASSWORD,
  });
  if (protectedState.structureProtected !== true) {
    throw new Error(`${label}: 结构保护未生效: ${JSON.stringify(protectedState)}`);
  }

  let deleteError: string | undefined;
  try {
    await bridge.sheetOperation("delete", "S3");
    throw new Error(`${label}: 结构保护下删除应失败，但成功了`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("结构保护下删除应失败")) throw error;
    deleteError = message;
  }
  if (!deleteError) throw new Error(`${label}: 未捕获结构保护删除错误`);

  const after = await getExcelDisplayAlerts();
  if (after.displayAlerts !== true) {
    throw new Error(
      `${label}: 结构保护删除失败后 DisplayAlerts 未恢复为 true: ${JSON.stringify(after)}`,
    );
  }
  await assertSheetNames(bridge, ["S1", "S3"], `${label}:sheet-remains`);

  const unprotected = await setExcelStructureProtected({
    protected: false,
    password: STRUCTURE_PASSWORD,
  });
  if (unprotected.structureProtected !== false) {
    throw new Error(`${label}: 解除结构保护失败: ${JSON.stringify(unprotected)}`);
  }

  return {
    original: true,
    displayAlertsAfter: after.displayAlerts,
    deleteError,
    sheets: ["S1", "S3"],
  };
}

async function caseLastVisibleFailRestore(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>> {
  await assertSheetNames(bridge, ["S1"], `${label}:before`);
  const set = await setExcelDisplayAlerts(true);
  if (set.displayAlerts !== true) {
    throw new Error(`${label}: 无法将 DisplayAlerts 设为 true: ${JSON.stringify(set)}`);
  }

  let deleteError: string | undefined;
  try {
    await bridge.sheetOperation("delete", "S1");
    throw new Error(`${label}: 删除最后可见工作表应失败，但成功了`);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("删除最后可见工作表应失败")) throw error;
    deleteError = message;
  }
  if (!deleteError) throw new Error(`${label}: 未捕获最后可见表删除错误`);

  const after = await getExcelDisplayAlerts();
  if (after.displayAlerts !== true) {
    throw new Error(
      `${label}: 最后可见表删除失败后 DisplayAlerts 未恢复为 true: ${JSON.stringify(after)}`,
    );
  }
  await assertSheetNames(bridge, ["S1"], `${label}:sheet-remains`);
  return {
    original: true,
    displayAlertsAfter: after.displayAlerts,
    deleteError,
    sheets: ["S1"],
  };
}

async function attachToFixtureExcel(bridge: DotNetExcelBridge, label: string): Promise<void> {
  const status = await bridge.selectHost("excel");
  if (!status.connected) {
    throw new Error(`${label}: Excel 未连接（openFixtures 应已打开独立进程）`);
  }
  const probe = await bridge.readRange("S1", "A1", "none");
  if (!probe.values || probe.values.length === 0) {
    throw new Error(`${label}: 活动工作簿不可读 ${JSON.stringify(probe)}`);
  }
}

async function assertSheetNames(
  bridge: DotNetExcelBridge,
  expected: string[],
  label: string,
): Promise<void> {
  const inspect = (await bridge.inspectWorkbook()) as {
    sheets?: Array<{ name?: string }>;
  };
  const actual = (inspect.sheets || [])
    .map((sheet) => String(sheet.name || ""))
    .filter((name) => name.length > 0)
    .sort();
  const sortedExpected = [...expected].sort();
  if (JSON.stringify(actual) !== JSON.stringify(sortedExpected)) {
    throw new Error(
      `${label}: 工作表列表不符 actual=${JSON.stringify(actual)} expected=${JSON.stringify(sortedExpected)}`,
    );
  }
}

async function assertNoHostStarted(
  microsoftBefore: number[],
  wpsAllBefore: number[],
  label: string,
): Promise<void> {
  const current = await listOfficeSmokeProcesses();
  const microsoftLeaked = current.microsoft.filter((id) => !microsoftBefore.includes(id));
  const wpsAll = current.wpsAll ?? current.wpsVisible;
  const wpsLeaked = wpsAll.filter((id) => !wpsAllBefore.includes(id));
  if (microsoftLeaked.length > 0 || wpsLeaked.length > 0) {
    throw new Error(
      `${label}: Open XML 夹具不应启动宿主进程 microsoft=${microsoftLeaked.join(",")} wps=${wpsLeaked.join(",")}`,
    );
  }
}

async function assertProcessesStillRunning(processIds: number[], label: string): Promise<void> {
  if (processIds.length === 0) return;
  const current = await runningOfficeSmokeProcesses(processIds);
  const missing = processIds.filter((id) => !current.includes(id));
  if (missing.length > 0) throw new Error(`${label}被意外关闭: ${missing.join(", ")}`);
}

async function assertOwnedStopped(ownedIds: number[]): Promise<void> {
  if (ownedIds.length === 0) return;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const running = await runningOfficeSmokeProcesses(ownedIds);
    if (running.length === 0) return;
    await sleep(250);
  }
  const still = await runningOfficeSmokeProcesses(ownedIds);
  if (still.length > 0) throw new Error(`冒烟登记进程未退出: ${still.join(", ")}`);
}

async function assertNoUnexpectedMicrosoft(before: number[]): Promise<void> {
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const current = (await listOfficeSmokeProcesses()).microsoft;
    const leaked = current.filter((id) => !before.includes(id));
    if (leaked.length === 0) return;
    await sleep(250);
  }
  const leaked = (await listOfficeSmokeProcesses()).microsoft.filter((id) => !before.includes(id));
  if (leaked.length > 0) throw new Error(`遗留未登记 Microsoft Office 进程: ${leaked.join(", ")}`);
}

function ownedIdsFrom(owned: OwnedExcel): number[] {
  return [...owned.excel, ...owned.word, ...owned.presentation];
}

function parseHostMode(raw: string | undefined): HostMode {
  const value = (raw || "excel").trim().toLowerCase();
  if (value === "excel" || value === "wps" || value === "both") return value;
  throw new Error(`WENGGE_EXCEL_DISPLAY_ALERTS_HOST 只能是 excel|wps|both，收到: ${raw}`);
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

void main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? error.stack || error.message : String(error)}\n`,
  );
  process.exitCode = 1;
});
