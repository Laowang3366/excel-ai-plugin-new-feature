import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DotNetExcelBridge } from "../electron/agent/officeWorker/dotNetExcelBridge";
import {
  closeOfficeFixtures,
  disposeOfficeWorker,
  applyExcelAdvancedAction,
  listOfficeSmokeProcesses,
  openOfficeFixtures,
  runningOfficeSmokeProcesses,
} from "./officeWorkerSmokeHelpers";

type HostMode = "excel" | "wps" | "both";

type OwnedExcel = { excel: number[]; word: number[]; presentation: number[] };

async function main(): Promise<void> {
  process.env.WENGGE_OFFICE_SMOKE = "1";
  const hostMode = parseHostMode(process.env.WENGGE_EXCEL_DYNAMIC_ARRAY_HOST);
  const summary: Record<string, unknown> = { hostMode, scenarios: [] as unknown[] };
  let tempDir: string | undefined;

  try {
    // listOfficeSmokeProcesses starts the Worker; keep all discovery/refusal under cleanup.
    const microsoftBefore = (await listOfficeSmokeProcesses()).microsoft;
    const wpsBefore = (await listOfficeSmokeProcesses()).wpsVisible;

    // selectHost resolves active COM via ROT; preexisting Excel can be modified by accident.
    if ((hostMode === "excel" || hostMode === "both") && microsoftBefore.length > 0) {
      throw new Error(
        `Excel 动态数组冒烟拒绝在已有 Microsoft Office 进程上运行（pids: ${microsoftBefore.join(", ")}）。请使用无用户 Office 窗口的隔离 Runner。`,
      );
    }

    tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-excel-dynamic-array-"));
    const filePath = path.join(tempDir, "dynamic-array.xlsx");
    await createFixture(filePath);
    if (hostMode === "excel" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(await runExcel365Matrix(filePath, microsoftBefore));
    }
    if (hostMode === "wps" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(await runWpsFormula2Capability(filePath, wpsBefore));
    }
    await assertProcessesStillRunning(microsoftBefore, "任务前 Microsoft Office 进程");
    await assertProcessesStillRunning(wpsBefore, "任务前可见 WPS 进程");
  } finally {
    await disposeOfficeWorker();
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
    }
  }
  // Only after scenarios, ownership checks, and cleanup all succeed.
  process.stdout.write(`${JSON.stringify({ ok: true, ...summary }, null, 2)}\n`);
}

async function createFixture(filePath: string): Promise<void> {
  const result = await applyExcelAdvancedAction({
    operation: "createWorkbook",
    filePath,
    params: {
      sheetNames: ["Sheet1"],
      values: [
        ["Name", "Score", "MarkerB", "MarkerC"],
        ["Alice", 90, "keep-b", "keep-c"],
        ["Bob", 80, "", ""],
        ["Carol", 70, "", ""],
      ],
    },
  });
  if (result.status !== "done") {
    throw new Error(`创建动态数组夹具失败: ${result.error || result.summary}`);
  }
}

async function runExcel365Matrix(
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
    await attachToFixtureExcel(bridge, "Excel 365 首次 openFixtures");

    const spillWrite = await bridge.writeRange("Sheet1", "F1", [["=SEQUENCE(3)"]]);
    if ((spillWrite.dynamicCells ?? 0) < 1) {
      throw new Error(`Excel SEQUENCE 未走动态数组写入: ${JSON.stringify(spillWrite)}`);
    }
    const spill = await bridge.readRange("Sheet1", "F1", "spill");
    assertSequenceSpill(spill.values, "Excel 365 表达式 spill");

    const beforeRollback = await bridge.readRange("Sheet1", "C2:D2", "none");
    const overlong = `=${"1+".repeat(9000)}1`;
    let rollbackError: string | undefined;
    try {
      await bridge.writeRange("Sheet1", "C2:D2", [["=A2", overlong]]);
      throw new Error("多公式第二项应失败，但写入成功");
    } catch (error) {
      rollbackError = error instanceof Error ? error.message : String(error);
    }
    const afterRollback = await bridge.readRange("Sheet1", "C2:D2", "none");
    assertMatrixEqual(afterRollback.values, beforeRollback.values, "多公式失败后整区回滚");

    const saved = await bridge.saveWorkbook();
    if (!saved.success) throw new Error(saved.error || "保存工作簿失败");
    const firstOwned = ownedIdsFrom(owned);
    await closeOfficeFixtures(owned);
    owned = undefined;
    await assertOwnedStopped(firstOwned);

    owned = await openOfficeFixtures({
      excelPaths: [filePath],
      wordPaths: [],
      presentationPaths: [],
    });
    await attachToFixtureExcel(bridge, "Excel 365 重开 openFixtures");
    const spillAfterReopen = await bridge.readRange("Sheet1", "F1", "spill");
    assertSequenceSpill(spillAfterReopen.values, "保存关闭重开后 spill");

    const secondOwned = ownedIdsFrom(owned);
    await closeOfficeFixtures(owned);
    owned = undefined;
    await assertOwnedStopped(secondOwned);
    await assertNoUnexpectedMicrosoft(microsoftBefore);

    return {
      host: "excel",
      spill: true,
      multiFormulaRollback: true,
      rollbackError,
      saveReopenSpill: true,
      formula2: true,
    };
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

/** openFixtures already opened the file; only select host and prove active workbook is readable. */
async function attachToFixtureExcel(bridge: DotNetExcelBridge, label: string): Promise<void> {
  const status = await bridge.selectHost("excel");
  if (!status.connected) {
    throw new Error(`${label}: Excel 未连接（openFixtures 应已打开独立进程）`);
  }
  const probe = await bridge.readRange("Sheet1", "A1", "none");
  if (!probe.values || probe.values.length === 0) {
    throw new Error(`${label}: 活动工作簿不可读 ${JSON.stringify(probe)}`);
  }
}

/**
 * WPS real-matrix path: must connect, open, write Formula2 SEQUENCE, and produce ordered spill.
 * Soft "capability unsupported" green results are not allowed.
 * Cleanup is disposeOfficeWorker only — never kill by PID delta.
 */
async function runWpsFormula2Capability(
  filePath: string,
  wpsBefore: number[],
): Promise<Record<string, unknown>> {
  if (wpsBefore.length > 0) {
    throw new Error(
      `WPS 动态数组冒烟拒绝附加已有可见 WPS 进程（pids: ${wpsBefore.join(", ")}）。请使用无用户 WPS 窗口的隔离 Runner。`,
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

  try {
    const write = await bridge.writeRange("Sheet1", "H1", [["=SEQUENCE(2)"]]);
    if ((write.dynamicCells ?? 0) < 1 && (write.written ?? 0) < 1) {
      throw new Error(`WPS SEQUENCE 写入失败: ${JSON.stringify(write)}`);
    }
    const spill = await bridge.readRange("Sheet1", "H1", "spill");
    assertSequenceSpill(spill.values, "WPS Formula2 SEQUENCE spill");
    const saved = await bridge.saveWorkbook();
    if (!saved.success) throw new Error(saved.error || "WPS 保存失败");

    return {
      host: "wps",
      formula2Supported: true,
      capability: "formula2_spill_ok",
      write,
      spillValues: spill.values,
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`WPS Formula2/spill 实机矩阵失败: ${message}`);
  } finally {
    // Ownership is Worker GetOrCreate/Create + disposeOfficeWorker (caller finally).
    // Never process.kill by visible-PID delta. Preexisting PIDs are asserted in main().
    await assertProcessesStillRunning(wpsBefore, "任务前可见 WPS 进程");
  }
}

function assertSequenceSpill(values: unknown[][] | undefined, label: string): void {
  if (!looksLikeSequence(values)) {
    throw new Error(`${label} 未得到 SEQUENCE 溢出结果: ${JSON.stringify(values)}`);
  }
}

function looksLikeSequence(values: unknown[][] | undefined): boolean {
  if (!values || values.length < 2) return false;
  const flat = values.map((row) => Number(Array.isArray(row) ? row[0] : row));
  if (flat.some((n) => !Number.isFinite(n))) return false;
  if (flat[0] !== 1) return false;
  for (let i = 1; i < flat.length; i += 1) {
    if (flat[i] !== flat[i - 1] + 1) return false;
  }
  return true;
}

function assertMatrixEqual(
  actual: unknown[][] | undefined,
  expected: unknown[][] | undefined,
  label: string,
): void {
  if (JSON.stringify(actual ?? null) !== JSON.stringify(expected ?? null)) {
    throw new Error(
      `${label} 失败: actual=${JSON.stringify(actual)} expected=${JSON.stringify(expected)}`,
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
  throw new Error(`WENGGE_EXCEL_DYNAMIC_ARRAY_HOST 只能是 excel|wps|both，收到: ${raw}`);
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
