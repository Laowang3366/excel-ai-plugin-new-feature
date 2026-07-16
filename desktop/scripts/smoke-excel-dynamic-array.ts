import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DotNetExcelBridge } from "../electron/agent/officeWorker/dotNetExcelBridge";
import { OfficeWorkerError } from "../electron/agent/officeWorker/officeWorkerClient";
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
    assertSequenceSpill(spill.values, "Excel 365 SEQUENCE spill");

    // H-08: four expression-type dynamic arrays on non-overlapping targets (source B2:B4).
    const expressionResults = await writeAndVerifyExpressionSpills(bridge, "Excel 365");

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

    // Legacy CSE: multi-cell FormulaArray via production writeRange + legacyCse:true.
    const legacyCse = await writeAndVerifyLegacyCse(bridge, "Excel 365");

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
    assertSequenceSpill(spillAfterReopen.values, "保存关闭重开后 SEQUENCE spill");
    const expressionAfterReopen = await verifyExpressionSpillsAfterReopen(bridge, "Excel 365 重开");

    const secondOwned = ownedIdsFrom(owned);
    await closeOfficeFixtures(owned);
    owned = undefined;
    await assertOwnedStopped(secondOwned);
    await assertNoUnexpectedMicrosoft(microsoftBefore);

    return {
      host: "excel",
      spill: true,
      expressionSpills: expressionResults,
      expressionSpillsAfterReopen: expressionAfterReopen,
      multiFormulaRollback: true,
      rollbackError,
      legacyCse,
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

  // Task-owned visible WPS PIDs only (exclude preexisting); used to prove dispose actually exited them.
  const wpsAfterOpen = (await listOfficeSmokeProcesses()).wpsVisible;
  const ownedWpsPids = wpsAfterOpen.filter((id) => !wpsBefore.includes(id));
  if (ownedWpsPids.length === 0) {
    throw new Error("WPS openWorkbook 后未观察到本任务创建的可见 WPS 进程");
  }

  try {
    const write = await bridge.writeRange("Sheet1", "H1", [["=SEQUENCE(2)"]]);
    if ((write.dynamicCells ?? 0) < 1 && (write.written ?? 0) < 1) {
      throw new Error(`WPS SEQUENCE 写入失败: ${JSON.stringify(write)}`);
    }
    const spill = await bridge.readRange("Sheet1", "H1", "spill");
    assertSequenceSpill(spill.values, "WPS Formula2 SEQUENCE spill");

    // H-07: same multi-formula path as Excel — first cell succeeds, second overlong fails; full region must restore.
    const beforeRollback = await bridge.readRange("Sheet1", "C2:D2", "none");
    const overlong = `=${"1+".repeat(9000)}1`;
    let rollbackError: string | undefined;
    try {
      await bridge.writeRange("Sheet1", "C2:D2", [["=A2", overlong]]);
      throw new Error("WPS 多公式第二项应失败，但写入成功");
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes("WPS 多公式第二项应失败")) throw error;
      rollbackError = message;
    }
    if (!rollbackError) {
      throw new Error("WPS 多公式第二项失败场景未捕获错误");
    }
    const afterRollback = await bridge.readRange("Sheet1", "C2:D2", "none");
    assertMatrixEqual(afterRollback.values, beforeRollback.values, "WPS 多公式失败后整区回滚");

    // Same legacyCse assertions as Excel; structured capability if FormulaArray unsupported.
    const legacyCse = await writeAndVerifyLegacyCse(bridge, "WPS");

    const saved = await bridge.saveWorkbook();
    if (!saved.success) throw new Error(saved.error || "WPS 保存失败");

    // H-08: dispose Worker, wait for task-owned WPS PIDs to exit, then reopen on a fresh client.
    await disposeOfficeWorker();
    await assertOwnedStopped(ownedWpsPids);
    const reopenBridge = new DotNetExcelBridge();
    const reselected = await reopenBridge.selectHost("wps");
    const reopened = await reopenBridge.openWorkbook(filePath).catch((error: unknown) => ({
      success: false as const,
      error: error instanceof Error ? error.message : String(error),
    }));
    if (!reopened.success) {
      throw new Error(
        `WPS 保存关闭后重开失败: ${reopened.error || `selectHost connected=${reselected.connected}`}`,
      );
    }
    const spillAfterReopen = await reopenBridge.readRange("Sheet1", "H1", "spill");
    assertSequenceSpill(spillAfterReopen.values, "WPS 保存关闭重开后 SEQUENCE spill");

    return {
      host: "wps",
      formula2Supported: true,
      capability: "formula2_spill_ok",
      multiFormulaRollback: true,
      legacyCse,
      saveCloseReopenSpill: true,
      ownedWpsPids,
      rollbackError,
      write,
      spillValues: spill.values,
      spillValuesAfterReopen: spillAfterReopen.values,
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

/**
 * Multi-cell traditional CSE via production writeRange + legacyCse:true.
 * Target N1:N3 does not overlap SEQUENCE (F1), expressions (J/K/L), or rollback (C2:D2).
 * Source B2:B4 = 90,80,70.
 */
async function writeAndVerifyLegacyCse(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>> {
  const formula = "=B2:B4";
  const range = "N1:N3";
  try {
    const write = await bridge.writeRange("Sheet1", range, [[formula]], { legacyCse: true });
    if (write.written !== 1 || write.arrayCells !== 1 || write.plainCells !== 0) {
      throw new Error(`${label} legacyCse 计数错误: ${JSON.stringify(write)} formula=${formula}`);
    }
    if (write.dynamicCells !== 0) {
      throw new Error(`${label} legacyCse 不应计为 dynamicCells: ${JSON.stringify(write)}`);
    }
    const currentArray = await bridge.readRange("Sheet1", "N1", "currentArray");
    assertMatrixEqual(currentArray.values, [[90], [80], [70]], `${label} CSE 值 N1:N3`);
    if (currentArray.address !== range) {
      throw new Error(`${label} CSE currentArray 范围错误: ${JSON.stringify(currentArray)}`);
    }
    const context = (await bridge.getFormulaContext("Sheet1", "N1")) as {
      formulas?: Array<{ formula?: string; address?: string }>;
    };
    const formulas = context.formulas || [];
    if (formulas.length === 0) {
      throw new Error(`${label} CSE: formula.context 未返回 N1 公式`);
    }
    const hostFormula = String(formulas[0]?.formula || "");
    if (!hostFormula.includes("B2") || hostFormula.includes("@")) {
      throw new Error(`${label} CSE: 读回公式不像 FormulaArray（got ${hostFormula}）`);
    }

    const rollbackRange = "O1:O3";
    await bridge.writeRange("Sheet1", "O1", [["keep-1"], ["keep-2"], ["keep-3"]]);
    const beforeRollback = await bridge.readRange("Sheet1", rollbackRange, "none");
    let rollbackErrorCode: string | undefined;
    try {
      const overlong = `=${"1+".repeat(9000)}1`;
      await bridge.writeRange("Sheet1", rollbackRange, [[overlong]], { legacyCse: true });
      throw new Error(`${label} CSE 超长公式应失败，但写入成功`);
    } catch (error) {
      if (!(error instanceof OfficeWorkerError)) throw error;
      rollbackErrorCode = error.code;
    }
    const afterRollback = await bridge.readRange("Sheet1", rollbackRange, "none");
    assertMatrixEqual(afterRollback.values, beforeRollback.values, `${label} CSE 失败后整区回滚`);

    return {
      status: "supported",
      formula,
      range,
      arrayCells: write.arrayCells,
      formulaArray: true,
      hostFormula,
      values: currentArray.values,
      rollbackErrorCode,
      rollbackVerified: true,
    };
  } catch (error) {
    if (error instanceof OfficeWorkerError && error.code === "legacy_array_unsupported") {
      return {
        status: "unsupported",
        code: error.code,
        formula,
        range,
        error: error.message,
      };
    }
    throw error;
  }
}

/** Four H-08 expression-type spills; targets do not overlap SEQUENCE (F1) or rollback (C2:D2). */
const EXPRESSION_SPILLS = [
  {
    id: "range_ref",
    formula: "=B2:B4",
    anchor: "J1",
    expected: [[90], [80], [70]],
  },
  {
    id: "range_mul",
    formula: "=B2:B4*2",
    anchor: "K1",
    expected: [[180], [160], [140]],
  },
  {
    id: "range_if",
    formula: '=IF(B2:B4>85,B2:B4,"")',
    anchor: "L1",
    expected: [[90], [""], [""]],
  },
  {
    id: "transpose",
    formula: "=TRANSPOSE(B2:B4)",
    anchor: "J5",
    expected: [[90, 80, 70]],
  },
] as const;

async function writeAndVerifyExpressionSpills(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>[]> {
  const results: Record<string, unknown>[] = [];
  for (const item of EXPRESSION_SPILLS) {
    const write = await bridge.writeRange("Sheet1", item.anchor, [[item.formula]]);
    if ((write.dynamicCells ?? 0) < 1) {
      throw new Error(
        `${label} ${item.id} 未走动态数组写入: ${JSON.stringify(write)} formula=${item.formula}`,
      );
    }
    const spill = await bridge.readRange("Sheet1", item.anchor, "spill");
    assertSpillValues(spill.values, item.expected, `${label} ${item.id} spill`);
    const hostFormulas = await assertFormulaNotAtDegraded(
      bridge,
      "Sheet1",
      item.anchor,
      `${label} ${item.id}`,
    );
    results.push({
      id: item.id,
      formula: item.formula,
      anchor: item.anchor,
      dynamicCells: write.dynamicCells,
      spillValues: spill.values,
      hostFormulas,
    });
  }
  return results;
}

async function verifyExpressionSpillsAfterReopen(
  bridge: DotNetExcelBridge,
  label: string,
): Promise<Record<string, unknown>[]> {
  const results: Record<string, unknown>[] = [];
  for (const item of EXPRESSION_SPILLS) {
    const spill = await bridge.readRange("Sheet1", item.anchor, "spill");
    assertSpillValues(spill.values, item.expected, `${label} ${item.id} spill`);
    const hostFormulas = await assertFormulaNotAtDegraded(
      bridge,
      "Sheet1",
      item.anchor,
      `${label} ${item.id}`,
    );
    results.push({
      id: item.id,
      formula: item.formula,
      anchor: item.anchor,
      spillValues: spill.values,
      hostFormulas,
    });
  }
  return results;
}

/** Reject any host-read formula containing `@` (implicit intersection may appear mid-expression). */
async function assertFormulaNotAtDegraded(
  bridge: DotNetExcelBridge,
  sheetName: string,
  range: string,
  label: string,
): Promise<Array<{ address?: string; formula: string }>> {
  const context = (await bridge.getFormulaContext(sheetName, range)) as {
    formulas?: Array<{ formula?: string; address?: string }>;
  };
  const formulas = context.formulas || [];
  if (formulas.length === 0) {
    throw new Error(`${label}: formula.context 未返回公式 ${range}`);
  }
  const hostFormulas: Array<{ address?: string; formula: string }> = [];
  for (const entry of formulas) {
    const formula = String(entry.formula || "");
    hostFormulas.push({ address: entry.address, formula });
    if (formula.includes("@")) {
      throw new Error(`${label}: 公式含 @ 降级: ${formula} @ ${entry.address || range}`);
    }
  }
  return hostFormulas;
}

function assertSpillValues(
  actual: unknown[][] | undefined,
  expected: readonly (readonly unknown[])[],
  label: string,
): void {
  const normalizedActual = normalizeSpillMatrix(actual);
  const normalizedExpected = normalizeSpillMatrix(expected as unknown[][]);
  if (JSON.stringify(normalizedActual) !== JSON.stringify(normalizedExpected)) {
    throw new Error(
      `${label} 失败: actual=${JSON.stringify(actual)} expected=${JSON.stringify(expected)}`,
    );
  }
}

function normalizeSpillMatrix(values: unknown[][] | undefined): unknown[][] {
  if (!values) return [];
  return values.map((row) =>
    (Array.isArray(row) ? row : [row]).map((cell) => {
      if (cell === null || cell === undefined) return "";
      if (typeof cell === "number" && Number.isFinite(cell)) return cell;
      const text = String(cell).trim();
      if (text === "") return "";
      const asNumber = Number(text);
      return Number.isFinite(asNumber) && text !== "" ? asNumber : text;
    }),
  );
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
