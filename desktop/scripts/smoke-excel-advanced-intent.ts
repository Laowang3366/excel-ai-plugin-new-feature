import { createHash } from "node:crypto";
import { copyFile, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import type { OfficeActionBridge } from "../electron/agent/tools/contracts/office";
import { addOfficeExecutors } from "../electron/agent/tools/executors/officeExecutors";
import { createOfficeActionBridge } from "../electron/agent/tools/officeCore/officeActionAdapter";
import type {
  OfficeActionInput,
  OfficeActionResult,
} from "../electron/agent/tools/officeCore/types";
import type { ToolExecutor } from "../electron/agent/shared/types";
import {
  applyExcelAdvancedAction,
  closeOfficeFixtures,
  disposeOfficeWorker,
  DotNetOfficeActionBridge,
  DotNetOfficeDocumentBridge,
  listOfficeSmokeProcesses,
  openOfficeFixtures,
  runningOfficeSmokeProcesses,
} from "./officeWorkerSmokeHelpers";

type HostMode = "excel" | "wps" | "both";

const PIVOT_NAME = "M05SalesPivot";
const QUERY_NAME = "M05ExternalCsv";
const SLICER_NAME = "M05DeptSlicer";
const SLICER_FIELD = "Dept";

/** Only explicit Worker capability codes — never free-text COM messages. */
const STABLE_UNSUPPORTED_CODES = new Set([
  "power_query_unavailable",
  "unsupported_operation",
  "unsupported_host",
  "unsupported_app",
  "capability_unsupported",
]);

async function main(): Promise<void> {
  process.env.WENGGE_OFFICE_SMOKE = "1";
  const hostMode = parseHostMode(process.env.WENGGE_EXCEL_ADVANCED_INTENT_HOST);
  const summary: Record<string, unknown> = { hostMode, scenarios: [] as unknown[] };
  let tempDir: string | undefined;
  let microsoftBefore: number[] = [];
  let wpsVisibleBefore: number[] = [];
  let residualAfterDispose: string | undefined;
  let tempDirCleanupError: string | undefined;

  try {
    const processesBefore = await listOfficeSmokeProcesses();
    microsoftBefore = processesBefore.microsoft;
    wpsVisibleBefore = processesBefore.wpsVisible;

    if ((hostMode === "excel" || hostMode === "both") && microsoftBefore.length > 0) {
      throw new Error(
        `高级意图冒烟拒绝在已有 Microsoft Office 进程上运行（pids: ${microsoftBefore.join(", ")}）`,
      );
    }
    // Visible document hosts only — background wpscloudsvr/CEF helpers are not a reject condition.
    if ((hostMode === "wps" || hostMode === "both") && wpsVisibleBefore.length > 0) {
      throw new Error(
        `高级意图冒烟拒绝附加已有可见 WPS 文档窗口（pids: ${wpsVisibleBefore.join(", ")}）`,
      );
    }

    tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-excel-advanced-intent-"));
    const excelPath = path.join(tempDir, "advanced-intent-excel.xlsx");
    const wpsBasePath = path.join(tempDir, "advanced-intent-wps-base.xlsx");
    const wpsAdvancedPath = path.join(tempDir, "advanced-intent-wps-advanced.xlsx");
    const csvPath = path.join(tempDir, "sales.csv");
    await writeFile(csvPath, "Dept,Amount\nA,10\nB,20\nA,30\n", "utf8");
    await createWorkbookFixture(excelPath);
    await assertNoHostStarted(microsoftBefore, wpsVisibleBefore, "createFixture");
    await copyFile(excelPath, wpsBasePath);
    await copyFile(excelPath, wpsAdvancedPath);

    // Negative boundary: production office.action.apply → operationPolicy, no COM.
    (summary.scenarios as unknown[]).push(await runNegativeIntentBoundary(excelPath));

    if (hostMode === "excel" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(
        await runExcelPositiveMatrix(excelPath, csvPath, microsoftBefore),
      );
    }
    if (hostMode === "wps" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(
        await runWpsMatrix(wpsBasePath, wpsAdvancedPath, csvPath),
      );
    }

    // Same basename in three dirs: simultaneous open + FullName/instanceId binding (Excel only).
    if (hostMode === "excel" || hostMode === "both") {
      (summary.scenarios as unknown[]).push(
        await runExcelSameNameCollision(tempDir, microsoftBefore),
      );
    }

    await assertProcessesStillRunning(microsoftBefore, "任务前 Microsoft Office 进程");
    await assertProcessesStillRunning(wpsVisibleBefore, "任务前可见 WPS 进程");
  } finally {
    if (hostMode === "wps" || hostMode === "both") {
      try {
        const after = await listOfficeSmokeProcesses();
        const visibleLeaked = after.wpsVisible.filter((id) => !wpsVisibleBefore.includes(id));
        if (visibleLeaked.length > 0) {
          residualAfterDispose = `结束后仍有新增可见 WPS 文档窗口: ${visibleLeaked.join(", ")}`;
        }
      } catch (error) {
        residualAfterDispose = `无法验证可见 WPS 窗口: ${
          error instanceof Error ? error.message : String(error)
        }`;
      }
    }
    await disposeOfficeWorker();
    // Do not delete temp while residual visible WPS still holds fixture paths (avoids "file not found" tabs).
    if (tempDir && !residualAfterDispose) {
      try {
        await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
      } catch (error) {
        tempDirCleanupError = error instanceof Error ? error.message : String(error);
      }
    } else if (tempDir && residualAfterDispose) {
      process.stderr.write(`[office-smoke] keep tempDir (visible residual): ${tempDir}\n`);
    }
  }

  if (residualAfterDispose) {
    throw new Error(residualAfterDispose);
  }
  if (tempDirCleanupError) {
    throw new Error(`tempDir 删除失败: ${tempDirCleanupError}`);
  }

  process.stdout.write(`${JSON.stringify({ ok: true, ...summary }, null, 2)}\n`);
}

async function createWorkbookFixture(filePath: string): Promise<void> {
  const result = await applyExcelAdvancedAction({
    operation: "createWorkbook",
    filePath,
    params: {
      sheetNames: ["Sheet1"],
      values: [
        ["Dept", "Amount"],
        ["A", 10],
        ["B", 20],
        ["A", 30],
      ],
    },
  });
  if (result.status !== "done") {
    throw new Error(`创建高级意图夹具失败: ${result.error || result.summary}`);
  }
}

/**
 * Production path: office.action.apply → officeAdvancedOperationError before bridge.
 * Counting wrapper proves COM/Worker is never reached; SHA-256 proves file unchanged.
 */
async function runNegativeIntentBoundary(filePath: string): Promise<Record<string, unknown>> {
  const { apply, getBridgeCalls } = createApplyExecutor();
  const beforeHash = await sha256File(filePath);
  const cases: Record<string, unknown>[] = [];

  const rejects: Array<{
    id: string;
    operation: string;
    action: string;
    target?: string;
    params: Record<string, unknown>;
    expect: string;
  }> = [
    {
      id: "pq_missing_advancedIntent",
      operation: "createPowerQuery",
      action: "edit",
      params: {
        name: "BadQuery",
        sourceKind: "external",
        mFormula: "let Source = 1 in Source",
        loadMode: "connectionOnly",
      },
      expect: "refreshable-etl",
    },
    {
      id: "pq_wrong_advancedIntent",
      operation: "createPowerQuery",
      action: "edit",
      params: {
        advancedIntent: "interactive-pivot",
        name: "BadQuery",
        sourceKind: "external",
        mFormula: "let Source = 1 in Source",
        loadMode: "connectionOnly",
      },
      expect: "refreshable-etl",
    },
    {
      id: "pq_missing_sourceKind",
      operation: "createPowerQuery",
      action: "edit",
      params: {
        advancedIntent: "refreshable-etl",
        name: "BadQuery",
        mFormula: "let Source = 1 in Source",
        loadMode: "connectionOnly",
      },
      expect: "sourceKind",
    },
    {
      id: "pivot_missing_advancedIntent",
      operation: "createPivotTable",
      action: "insert",
      target: "range:Sheet1!A1:B4",
      params: { rowFields: ["Dept"], dataFields: [{ name: "Amount", function: "sum" }] },
      expect: "interactive-pivot",
    },
    {
      id: "pivot_wrong_advancedIntent",
      operation: "createPivotTable",
      action: "insert",
      target: "range:Sheet1!A1:B4",
      params: {
        advancedIntent: "refreshable-etl",
        rowFields: ["Dept"],
        dataFields: [{ name: "Amount", function: "sum" }],
      },
      expect: "interactive-pivot",
    },
    {
      id: "slicer_missing_advancedIntent",
      operation: "addSlicer",
      action: "insert",
      target: "range:Sheet1!A1",
      params: { pivotName: PIVOT_NAME, field: SLICER_FIELD },
      expect: "interactive-pivot",
    },
    {
      id: "slicer_wrong_advancedIntent",
      operation: "addSlicer",
      action: "insert",
      target: "range:Sheet1!A1",
      params: {
        advancedIntent: "refreshable-etl",
        pivotName: PIVOT_NAME,
        field: SLICER_FIELD,
      },
      expect: "interactive-pivot",
    },
  ];

  for (const item of rejects) {
    const callsBefore = getBridgeCalls();
    const result = await apply.execute({
      app: "excel",
      action: item.action,
      operation: item.operation,
      filePath,
      target: item.target,
      params: item.params,
    });
    if (result.success) {
      throw new Error(`${item.id}: 应拒绝但成功`);
    }
    const error = String(result.error || "");
    if (!error.includes(item.expect)) {
      throw new Error(`${item.id}: 拒绝消息未含 ${item.expect}: ${error}`);
    }
    if (getBridgeCalls() !== callsBefore) {
      throw new Error(`${item.id}: bridge 调用计数增加（应在 COM 前拒绝）`);
    }
    const afterHash = await sha256File(filePath);
    if (afterHash !== beforeHash) {
      throw new Error(`${item.id}: 文件 SHA-256 变化`);
    }
    cases.push({ id: item.id, rejected: true, error, bridgeCalls: getBridgeCalls() });
  }

  return {
    host: "policy",
    path: "office.action.apply",
    sha256Unchanged: true,
    bridgeCalls: getBridgeCalls(),
    cases,
  };
}

async function runExcelPositiveMatrix(
  filePath: string,
  csvPath: string,
  microsoftBefore: number[],
): Promise<Record<string, unknown>> {
  const { apply, getBridgeCalls } = createApplyExecutor();
  const host = "excel";

  {
    const base = await applyBasicWrite(apply, filePath, host);
    if (!base.success) {
      throw new Error(`Excel 基础写值失败: ${base.error}`);
    }

    const pq = await applyPowerQuery(apply, filePath, csvPath, host);
    if (!pq.success) {
      throw new Error(`Excel createPowerQuery 失败: ${pq.error}`);
    }
    const inspect = await apply.execute({
      app: "excel",
      action: "inspect",
      operation: "inspectPowerQueries",
      filePath,
      params: { host },
    });
    if (!inspect.success) {
      throw new Error(`Excel inspectPowerQueries 失败: ${inspect.error}`);
    }
    assertPowerQueryInspect(inspect.data, csvPath);

    const pivot = await applyPivot(apply, filePath, host);
    if (!pivot.success) {
      throw new Error(`Excel createPivotTable 失败: ${pivot.error}`);
    }
    assertPivotResult(pivot.data);

    const slicer = await applySlicer(apply, filePath, host);
    if (!slicer.success) {
      throw new Error(`Excel addSlicer 失败: ${slicer.error}`);
    }
    assertSlicerResult(slicer.data);

    if (getBridgeCalls() < 1) {
      throw new Error("Excel 正向矩阵未调用 bridge");
    }

    await assertOwnedMicrosoftStopped(microsoftBefore);

    return {
      host: "excel",
      capability: "full",
      baseWrite: true,
      powerQuery: true,
      pivot: true,
      slicer: true,
      bridgeCalls: getBridgeCalls(),
    };
  }
}

/** Three directories, identical basename collision.xlsx: independent Excel PIDs + FullName/instanceId. */
async function runExcelSameNameCollision(
  tempDir: string,
  microsoftBefore: number[],
): Promise<Record<string, unknown>> {
  const dirs = ["A", "B", "C"] as const;
  const paths: string[] = [];
  const stage = (name: string) => {
    process.stderr.write(`[office-smoke] collision:${name}\n`);
  };

  stage("createFixtures.start");
  for (const dir of dirs) {
    const folder = path.join(tempDir, "collision-excel", dir);
    await mkdirp(folder);
    const filePath = path.join(folder, "collision.xlsx");
    const created = await applyExcelAdvancedAction({
      operation: "createWorkbook",
      filePath,
      params: { sheetNames: ["Sheet1"], values: [["Dir", dir]] },
    });
    if (created.status !== "done") {
      throw new Error(`创建同名夹具失败 ${filePath}: ${created.error || created.summary}`);
    }
    paths.push(filePath);
  }
  stage("createFixtures.done");

  let owned: { excel: number[]; word: number[]; presentation: number[] } | undefined;
  try {
    stage("openFixtures.start");
    owned = await openOfficeFixtures({
      excelPaths: paths,
      wordPaths: [],
      presentationPaths: [],
    });
    stage(`openFixtures.done:pids=${owned.excel.join(",")}`);
    if (owned.excel.length !== 3) {
      throw new Error(
        `openOfficeFixtures 应返回 3 个 owned Excel PID，实际 ${owned.excel.length}: ${JSON.stringify(owned)}`,
      );
    }

    const bridge = new DotNetOfficeDocumentBridge();
    stage("listDocuments.start");
    const documents = (await bridge.listDocuments("excel")).filter((doc) =>
      paths.some((p) => samePath(p, doc.fullName || "")),
    );
    stage(`listDocuments.done:count=${documents.length}`);
    if (documents.length !== 3) {
      throw new Error(
        `listDocuments 应按完整 FullName 返回 3 项，实际 ${documents.length}: ${JSON.stringify(documents)}`,
      );
    }
    const instanceIds = new Set(documents.map((d) => d.instanceId));
    if (instanceIds.size !== 3) {
      throw new Error(`三个 instanceId 必须互不相同: ${JSON.stringify(documents)}`);
    }

    for (const filePath of paths) {
      const match = documents.find((d) => samePath(filePath, d.fullName || ""));
      if (!match?.instanceId) {
        throw new Error(`未按 FullName 找到实例: ${filePath}`);
      }
      stage(`activate.start:${path.basename(path.dirname(filePath))}:${match.instanceId}`);
      await bridge.activateDocument({
        app: "excel",
        filePath,
        instanceId: match.instanceId,
      });
      stage(`activate.done:${path.basename(path.dirname(filePath))}`);
    }

    const ownedIds = [...owned.excel, ...owned.word, ...owned.presentation];
    stage("closeFixtures.start");
    await closeOfficeFixtures(owned);
    owned = undefined;
    stage("closeFixtures.done");
    await assertOwnedStopped(ownedIds);
    await assertOwnedMicrosoftStopped(microsoftBefore);
    stage("assertStopped.done");
    return {
      host: "excel",
      scenario: "same-name-collision",
      ownedExcelPids: 3,
      fullNameCount: 3,
      distinctInstanceIds: 3,
      activateByInstanceIdAndFullPath: true,
      files: paths.map((p) => path.relative(tempDir, p)),
    };
  } finally {
    if (owned) {
      try {
        await closeOfficeFixtures(owned);
      } catch {
        /* fixture cleanup only */
      }
    }
  }
}

async function mkdirp(dir: string): Promise<void> {
  const { mkdir } = await import("node:fs/promises");
  await mkdir(dir, { recursive: true });
}

function samePath(left: string, right: string): boolean {
  try {
    return path.resolve(left).toLowerCase() === path.resolve(right).toLowerCase();
  } catch {
    return left.toLowerCase() === right.toLowerCase();
  }
}

/** WPS base host proof + advanced capability classification. */
async function runWpsMatrix(
  baseFilePath: string,
  advancedFilePath: string,
  csvPath: string,
): Promise<Record<string, unknown>> {
  const host = "wps";
  const { apply, getBridgeCalls } = createApplyExecutor();
  const stage = (name: string) => {
    process.stderr.write(`[office-smoke] wps-matrix:${name}\n`);
  };

  stage("base.start");
  const base = await applyBasicWrite(apply, baseFilePath, host);
  stage(`base.done:success=${base.success}`);
  if (!base.success) {
    throw new Error(`WPS 基础写值失败（宿主不可用）: ${base.error}`);
  }

  stage("powerQuery.start");
  const powerQuery = await classifyAdvanced({
    operation: "createPowerQuery",
    run: () => applyPowerQuery(apply, advancedFilePath, csvPath, host),
    onSupported: async () => {
      stage("inspectPowerQueries.start");
      const inspect = await apply.execute({
        app: "excel",
        action: "inspect",
        operation: "inspectPowerQueries",
        filePath: advancedFilePath,
        params: { host },
      });
      stage(`inspectPowerQueries.done:success=${inspect.success}`);
      if (!inspect.success) {
        throw new Error(`WPS inspectPowerQueries 失败: ${inspect.error}`);
      }
      assertPowerQueryInspect(inspect.data, csvPath);
    },
  });
  stage(`powerQuery.done:status=${powerQuery.status}`);

  stage("pivot.start");
  const pivot = await classifyAdvanced({
    operation: "createPivotTable",
    run: () => applyPivot(apply, advancedFilePath, host),
    onSupported: async (data) => {
      assertPivotResult(data);
    },
  });
  stage(`pivot.done:status=${pivot.status}`);

  stage("slicer.start");
  const slicer = await classifyAdvanced({
    operation: "addSlicer",
    run: () => applySlicer(apply, advancedFilePath, host),
    onSupported: async (data) => {
      assertSlicerResult(data);
    },
  });
  stage(`slicer.done:status=${slicer.status}`);

  return {
    host: "wps",
    baseWrite: true,
    baseFixture: path.basename(baseFilePath),
    advancedFixture: path.basename(advancedFilePath),
    advanced: { powerQuery, pivot, slicer },
    bridgeCalls: getBridgeCalls(),
  };
}

/** Basic COM file-level op proving the selected host can open and mutate the workbook. */
async function applyBasicWrite(
  apply: ToolExecutor,
  filePath: string,
  host: "excel" | "wps",
): Promise<{ success: boolean; error?: string; data?: unknown }> {
  return apply.execute({
    app: "excel",
    action: "insert",
    operation: "manageWorkbookObject",
    filePath,
    params: {
      host,
      objectType: "worksheet",
      command: "add",
      name: host === "excel" ? "M05ExcelProof" : "M05WpsProof",
    },
  });
}

async function applyPowerQuery(
  apply: ToolExecutor,
  filePath: string,
  csvPath: string,
  host: "excel" | "wps",
): Promise<{ success: boolean; error?: string; data?: unknown }> {
  const formulaPath = csvPath.replace(/\\/g, "/");
  return apply.execute({
    app: "excel",
    action: "edit",
    operation: "createPowerQuery",
    filePath,
    params: {
      host,
      advancedIntent: "refreshable-etl",
      sourceKind: "external",
      name: QUERY_NAME,
      mFormula: `let Source = Csv.Document(File.Contents("${formulaPath}"),[Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]), #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]) in #"Promoted Headers"`,
      loadMode: "connectionOnly",
    },
  });
}

async function applyPivot(
  apply: ToolExecutor,
  filePath: string,
  host: "excel" | "wps",
): Promise<{ success: boolean; error?: string; data?: unknown }> {
  return apply.execute({
    app: "excel",
    action: "insert",
    operation: "createPivotTable",
    filePath,
    target: "range:Sheet1!A1:B4",
    params: {
      host,
      advancedIntent: "interactive-pivot",
      name: PIVOT_NAME,
      destination: "Sheet1!F3",
      rowFields: ["Dept"],
      dataFields: [{ name: "Amount", function: "sum" }],
    },
  });
}

async function applySlicer(
  apply: ToolExecutor,
  filePath: string,
  host: "excel" | "wps",
): Promise<{ success: boolean; error?: string; data?: unknown }> {
  return apply.execute({
    app: "excel",
    action: "insert",
    operation: "addSlicer",
    filePath,
    target: "range:Sheet1!A1",
    params: {
      host,
      advancedIntent: "interactive-pivot",
      pivotName: PIVOT_NAME,
      field: SLICER_FIELD,
      name: SLICER_NAME,
    },
  });
}

function createApplyExecutor(): {
  apply: ToolExecutor;
  getBridgeCalls: () => number;
} {
  let bridgeCalls = 0;
  const inner = new DotNetOfficeActionBridge();
  const countingCom: OfficeActionBridge = {
    executeAction: async (input: OfficeActionInput): Promise<OfficeActionResult> => {
      bridgeCalls += 1;
      return inner.executeAction(input);
    },
  };
  const adapter = createOfficeActionBridge({ officeComActionBridge: countingCom });
  const target = new Map<string, ToolExecutor>();
  addOfficeExecutors(target, { officeActionBridge: adapter });
  const apply = target.get("office.action.apply");
  if (!apply) throw new Error("office.action.apply 未注册");
  return { apply, getBridgeCalls: () => bridgeCalls };
}

function assertPowerQueryInspect(data: unknown, csvPath: string): void {
  const result = asRecord(data);
  const payload = asRecord(result.data ?? result);
  const queries = Array.isArray(payload.queries)
    ? payload.queries.map(asRecord)
    : Array.isArray(asRecord(payload.snapshot).queries)
      ? (asRecord(payload.snapshot).queries as unknown[]).map(asRecord)
      : [];
  const match = queries.find((q) => String(q.name || "") === QUERY_NAME);
  if (!match) {
    throw new Error(`inspectPowerQueries 未找到 ${QUERY_NAME}: ${JSON.stringify(payload)}`);
  }
  const formula = String(match.formula || "");
  const marker = csvPath.replace(/\\/g, "/");
  if (!formula.includes(path.basename(csvPath)) && !formula.includes(marker)) {
    throw new Error(`Power Query 公式未指向 CSV: ${formula}`);
  }
}

function unwrapActionResult(data: unknown): Record<string, unknown> {
  // office.action.apply returns { success, data: OfficeActionResult }
  const outer = asRecord(data);
  if (typeof outer.status === "string" && (outer.operation || outer.engine || outer.data)) {
    return outer;
  }
  return asRecord(outer.data ?? outer);
}

function assertPivotResult(data: unknown): void {
  const actionResult = unwrapActionResult(data);
  if (actionResult.status !== "done") {
    throw new Error(`透视表 status 非 done: ${JSON.stringify(actionResult)}`);
  }
  const payload = asRecord(actionResult.data);
  const readback = asRecord(payload.readback);
  const pivotName = String(readback.pivotName || readback.name || "");
  if (pivotName !== PIVOT_NAME) {
    throw new Error(
      `透视表 name 回读失败: expected=${PIVOT_NAME} actual=${pivotName} raw=${JSON.stringify(payload)}`,
    );
  }
  const rowFieldCount = Number(readback.rowFieldCount ?? 0);
  const dataFieldCount = Number(readback.dataFieldCount ?? 0);
  if (rowFieldCount < 1 || dataFieldCount < 1) {
    throw new Error(
      `透视表字段回读不足: rowFieldCount=${rowFieldCount} dataFieldCount=${dataFieldCount}`,
    );
  }
  const sourceAddress = String(readback.sourceAddress || "");
  const tableRange = String(readback.tableRange1 || "") || String(readback.destinationRange || "");
  if (!sourceAddress || !tableRange) {
    throw new Error(`透视表范围回读缺失: sourceAddress=${sourceAddress} tableRange=${tableRange}`);
  }
  const verification = asRecord(readback.verification);
  if (verification.ok !== true) {
    throw new Error(`透视表 verification.ok 非 true: ${JSON.stringify(readback)}`);
  }
}

function assertSlicerResult(data: unknown): void {
  const actionResult = unwrapActionResult(data);
  if (actionResult.status !== "done") {
    throw new Error(`切片器 status 非 done: ${JSON.stringify(actionResult)}`);
  }
  const payload = asRecord(actionResult.data);
  const slicerName = String(payload.slicerName || "");
  if (slicerName !== SLICER_NAME) {
    throw new Error(
      `切片器 name 回读失败: expected=${SLICER_NAME} actual=${slicerName} raw=${JSON.stringify(payload)}`,
    );
  }
}

/**
 * WPS advanced ops: supported → same object readback asserts; unsupported only via
 * structured Worker codes (power_query_unavailable / unsupported_operation / …).
 * Generic COM messages fail the smoke (Codex will then add minimal classification).
 */
async function classifyAdvanced(input: {
  operation: string;
  run: () => Promise<{ success: boolean; error?: string; data?: unknown }>;
  onSupported: (data: unknown) => Promise<void>;
}): Promise<Record<string, unknown>> {
  const result = await input.run();
  if (result.success) {
    await input.onSupported(result.data);
    return { operation: input.operation, capability: "supported", success: true };
  }

  const code = extractWorkerCode(result);
  if (code && STABLE_UNSUPPORTED_CODES.has(code)) {
    return {
      operation: input.operation,
      capability: "unsupported",
      success: false,
      code,
      error: result.error,
    };
  }
  throw new Error(
    `WPS ${input.operation} 失败且无结构化 capability code（code=${code ?? "none"}）: ${result.error || JSON.stringify(result.data)}`,
  );
}

function extractWorkerCode(result: { error?: string; data?: unknown }): string | undefined {
  const actionResult = unwrapActionResult(result.data);
  const nested = asRecord(actionResult.data);
  for (const candidate of [nested.code, actionResult.code, asRecord(result.data).code]) {
    if (typeof candidate === "string" && candidate.length > 0) return candidate;
  }
  // OfficeWorkerError message path: not a code. Do not parse free text.
  return undefined;
}

async function sha256File(filePath: string): Promise<string> {
  const bytes = await readFile(filePath);
  return createHash("sha256").update(bytes).digest("hex");
}

async function assertNoHostStarted(
  microsoftBefore: number[],
  wpsVisibleBefore: number[],
  label: string,
): Promise<void> {
  const current = await listOfficeSmokeProcesses();
  const microsoftLeaked = current.microsoft.filter((id) => !microsoftBefore.includes(id));
  const wpsVisibleLeaked = current.wpsVisible.filter((id) => !wpsVisibleBefore.includes(id));
  if (microsoftLeaked.length > 0 || wpsVisibleLeaked.length > 0) {
    throw new Error(
      `${label}: Open XML 夹具不应启动宿主 microsoft=${microsoftLeaked.join(",")} wpsVisible=${wpsVisibleLeaked.join(",")}`,
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
  for (let attempt = 0; attempt < 40; attempt += 1) {
    const running = await runningOfficeSmokeProcesses(ownedIds);
    if (running.length === 0) return;
    await sleep(250);
  }
  const still = await runningOfficeSmokeProcesses(ownedIds);
  if (still.length > 0) throw new Error(`冒烟登记进程未退出: ${still.join(", ")}`);
}

async function assertOwnedMicrosoftStopped(microsoftBefore: number[]): Promise<void> {
  for (let attempt = 0; attempt < 40; attempt += 1) {
    const current = (await listOfficeSmokeProcesses()).microsoft;
    const extra = current.filter((id) => !microsoftBefore.includes(id));
    if (extra.length === 0) return;
    await sleep(250);
  }
  const extra = (await listOfficeSmokeProcesses()).microsoft.filter(
    (id) => !microsoftBefore.includes(id),
  );
  if (extra.length > 0) {
    throw new Error(`Excel 正向矩阵后遗留进程: ${extra.join(", ")}`);
  }
}

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : {};
}

function parseHostMode(raw: string | undefined): HostMode {
  const value = (raw || "excel").trim().toLowerCase();
  if (value === "excel" || value === "wps" || value === "both") return value;
  throw new Error(`WENGGE_EXCEL_ADVANCED_INTENT_HOST 只能是 excel|wps|both，收到: ${raw}`);
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
