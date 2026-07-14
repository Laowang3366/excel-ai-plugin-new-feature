import { randomUUID } from "node:crypto";
import path from "node:path";

import type { OfficeActionBridge } from "../contracts/office";
import {
  beginOfficeTransaction,
  finalizeOfficeTransaction,
  getOfficeTransaction,
  listOfficeTransactionPaths,
  recordOfficeTransactionResult,
  undoOfficeTransaction,
  type OfficeTransactionRestoreFile,
  type OfficeTransactionRecord,
} from "./transactionJournal";
import type { OfficeActionInput, OfficeActionResult } from "./types";
import {
  executeWorkflowStepWithRetry,
  resolveWorkflowStep,
  resolveWorkflowVariables,
  shouldRunWorkflowStep,
  type OfficeWorkflowStepInput,
} from "./workflowStepExecution";
import {
  acquireOfficeWorkflowLock,
  getOfficeWorkflowRecord,
  isOfficeWorkflowCancellationRequested,
  listOfficeWorkflowRecords,
  saveOfficeWorkflowRecord,
  startOfficeWorkflowLockHeartbeat,
} from "./workflowRecordStore";
import { rollbackOfficeResults } from "./workflowRollback";

export { requestOfficeWorkflowCancellation } from "./workflowRecordStore";

export type OfficeWorkflowStatus = "running" | "paused" | "done" | "failed" | "cancelled";

export interface OfficeWorkflowStepRecord {
  step: number;
  id?: string;
  status: "pending" | "running" | "done" | "failed" | "skipped";
  startedAt?: string;
  completedAt?: string;
  attempts?: number;
  artifacts: string[];
  resolvedStep?: OfficeActionInput;
  result?: OfficeActionResult;
}

export interface OfficeWorkflowRecord {
  id: string;
  status: OfficeWorkflowStatus;
  createdAt: string;
  updatedAt: string;
  steps: OfficeWorkflowStepInput[];
  sourceSteps?: OfficeWorkflowStepInput[];
  stepRecords: OfficeWorkflowStepRecord[];
  completedSteps: number;
  nextStep: number;
  transactionId?: string;
  runToken?: string;
  leaseExpiresAt?: string;
  cancelRequested?: boolean;
  variables?: Record<string, unknown>;
  error?: string;
}

export interface OfficeWorkflowResult {
  status: "done" | "failed" | "paused" | "cancelled";
  workflowId?: string;
  transactionId?: string;
  completedSteps: number;
  nextStep?: number;
  failedStep?: number;
  results: OfficeActionResult[];
  stepRecords: OfficeWorkflowStepRecord[];
  rollback: Array<{ step: number; ok: boolean; summary: string }>;
  error?: string;
}

export interface OfficeWorkflowRunOptions {
  workflowRoot?: string;
  transactionRoot?: string;
  workflowId?: string;
  resume?: boolean;
  recoverRunning?: boolean;
  leaseMs?: number;
  failureMode?: "pause" | "rollback";
  cancellationMode?: "pause" | "rollback";
  variables?: Record<string, unknown>;
  prepareTransaction?: (filePaths: string[]) => Promise<unknown>;
  restoreTransaction?: (files: OfficeTransactionRestoreFile[]) => Promise<unknown>;
}

export async function runOfficeWorkflow(
  bridge: OfficeActionBridge,
  requestedSteps: OfficeWorkflowStepInput[],
  options: OfficeWorkflowRunOptions = {},
): Promise<OfficeWorkflowResult> {
  if (!options.workflowRoot) return runOfficeWorkflowUnlocked(bridge, requestedSteps, options);
  const workflowId = options.workflowId || randomUUID();
  let lock: Awaited<ReturnType<typeof acquireOfficeWorkflowLock>>;
  try { lock = await acquireOfficeWorkflowLock(options.workflowRoot, workflowId); }
  catch (error) { return failedResult(errorMessage(error)); }
  const stopHeartbeat = startOfficeWorkflowLockHeartbeat(lock);
  try {
    await lock.assertOwned();
    return await runOfficeWorkflowUnlocked(bridge, requestedSteps, { ...options, workflowId }, lock.assertOwned);
  } catch (error) {
    return failedResult(errorMessage(error));
  } finally {
    stopHeartbeat();
    await lock.release();
  }
}

async function runOfficeWorkflowUnlocked(
  bridge: OfficeActionBridge,
  requestedSteps: OfficeWorkflowStepInput[],
  options: OfficeWorkflowRunOptions,
  assertExecutionLease?: () => Promise<void>,
): Promise<OfficeWorkflowResult> {
  const resumed = options.resume === true;
  let record: OfficeWorkflowRecord | undefined;
  let steps = requestedSteps;
  let transaction: OfficeTransactionRecord | undefined;

  if (resumed) {
    if (!options.workflowRoot || !options.workflowId) return failedResult("继续工作流需要 workflowRoot 和 workflowId");
    try { record = await getOfficeWorkflowRecord(options.workflowRoot, options.workflowId); }
    catch (error) { return failedResult(errorMessage(error)); }
    const leaseExpired = !record.leaseExpiresAt || Date.parse(record.leaseExpiresAt) <= Date.now();
    const recoverRunning = record.status === "running" && (options.recoverRunning === true || leaseExpired);
    if (!["paused", "failed", "cancelled"].includes(record.status) && !recoverRunning) {
      return failedResult("只有暂停、失败、已取消或租约过期的工作流可以继续", record);
    }
    if (recoverRunning) {
      for (const stepRecord of record.stepRecords) {
        if (stepRecord.status === "running") stepRecord.status = "failed";
      }
      record.status = "paused";
    }
    if (requestedSteps.length > 0 && JSON.stringify(requestedSteps) !== JSON.stringify(record.sourceSteps || record.steps)) {
      return failedResult("继续工作流时提供的 steps 与原执行计划不一致", record);
    }
    steps = record.steps;
    if (options.transactionRoot && record.transactionId) {
      try { transaction = await getOfficeTransaction(options.transactionRoot, record.transactionId); }
      catch (error) { return failedResult(errorMessage(error), record); }
    }
  }

  if (!resumed && options.variables) {
    try { steps = resolveWorkflowVariables(steps, options.variables); }
    catch (error) { return failedResult(errorMessage(error)); }
  }

  const validationError = validateWorkflow(steps);
  if (validationError) return failedResult(validationError, record);

  if (!record) {
    const id = options.workflowId || randomUUID();
    const now = new Date().toISOString();
    record = {
      id,
      status: "running",
      createdAt: now,
      updatedAt: now,
      steps,
      sourceSteps: requestedSteps,
      stepRecords: steps.map((step, index) => ({ step: index + 1, id: step.id, status: "pending", artifacts: [] })),
      completedSteps: 0,
      nextStep: 1,
      variables: options.variables,
    };
    if (options.transactionRoot) {
      if (options.prepareTransaction) {
        try { await options.prepareTransaction(listOfficeTransactionPaths(steps)); }
        catch (error) { return failedResult(`准备已打开的 Office 文档失败: ${errorMessage(error)}`, record); }
      }
      transaction = await beginOfficeTransaction({ root: options.transactionRoot, steps, workflowId: id });
      record.transactionId = transaction.id;
    }
  }

  record.status = "running";
  record.runToken = randomUUID();
  record.leaseExpiresAt = leaseExpiry(options.leaseMs);
  record.cancelRequested = false;
  record.error = undefined;
  await saveWorkflowIfConfigured(options.workflowRoot, record);
  let index = Math.max(0, record.nextStep - 1);
  while (index < steps.length) {
    if (assertExecutionLease) await assertExecutionLease();
    if (await isOfficeWorkflowCancellationRequested(options.workflowRoot, record)) {
      const cancellationMode = options.cancellationMode || "pause";
      if (cancellationMode === "rollback") {
        const rollback = transaction && options.transactionRoot
          ? await rollbackTransaction(options.transactionRoot, transaction, index + 1, options)
          : await rollbackOfficeResults(bridge, steps, record.stepRecords.map((item) => item.result).filter(isResult));
        record.status = "cancelled";
        await saveWorkflowIfConfigured(options.workflowRoot, record);
        return resultFromRecord(record, "cancelled", rollback);
      }
      record.status = "cancelled";
      await saveWorkflowIfConfigured(options.workflowRoot, record);
      return resultFromRecord(record, "cancelled", []);
    }

    if (["done", "skipped"].includes(record.stepRecords[index].status)) {
      index++;
      continue;
    }
    const batch = workflowBatchIndexes(steps, index);
    const runnable: Array<{ index: number; resolved: OfficeActionInput }> = [];
    for (const stepIndex of batch) {
      const step = steps[stepIndex];
      const stepRecord = record.stepRecords[stepIndex];
      if (["done", "skipped"].includes(stepRecord.status)) continue;
      if (!shouldRunWorkflowStep(step.when, record.stepRecords)) {
        stepRecord.status = "skipped";
        stepRecord.completedAt = new Date().toISOString();
        continue;
      }
      try {
        stepRecord.resolvedStep = resolveWorkflowStep(step, record.stepRecords, record.variables);
        runnable.push({ index: stepIndex, resolved: { ...stepRecord.resolvedStep, transactionContext: "workflow" } });
        stepRecord.status = "running";
        stepRecord.startedAt = new Date().toISOString();
        stepRecord.completedAt = undefined;
        stepRecord.result = undefined;
        stepRecord.artifacts = [];
      } catch (error) {
        stepRecord.status = "failed";
        stepRecord.completedAt = new Date().toISOString();
        stepRecord.result = resolutionFailure(step, errorMessage(error));
      }
    }
    updateWorkflowProgress(record);
    record.leaseExpiresAt = leaseExpiry(options.leaseMs);
    await saveWorkflowIfConfigured(options.workflowRoot, record);

    const executions = await Promise.all(runnable.map(async (item) => ({
      ...item,
      execution: await executeWorkflowStepWithRetry(bridge, item.resolved, steps[item.index].retry),
    })));
    for (const item of executions) {
      const stepRecord = record.stepRecords[item.index];
      const { result, attempts } = item.execution;
      stepRecord.result = result;
      stepRecord.attempts = attempts;
      stepRecord.completedAt = new Date().toISOString();
      stepRecord.artifacts = collectStepArtifacts(result);
      stepRecord.status = result.status === "done" ? "done" : "failed";
      if (transaction && options.transactionRoot) {
        await recordOfficeTransactionResult(options.transactionRoot, transaction, result);
      }
    }
    updateWorkflowProgress(record);
    const failedIndex = batch.find((stepIndex) => record.stepRecords[stepIndex].status === "failed");
    if (failedIndex !== undefined) {
      const failedRecord = record.stepRecords[failedIndex];
      record.error = failedRecord.result?.error || failedRecord.result?.summary || "Office 工作流步骤失败";
      const failureMode = options.failureMode || (options.workflowRoot ? "pause" : "rollback");
      if (failureMode === "pause") {
        record.status = "paused";
        record.nextStep = failedIndex + 1;
        await saveWorkflowIfConfigured(options.workflowRoot, record);
        return resultFromRecord(record, "paused", [], failedIndex + 1);
      }
      const rollback = transaction && options.transactionRoot
        ? await rollbackTransaction(options.transactionRoot, transaction, failedIndex + 1, options)
        : await rollbackOfficeResults(bridge, steps, record.stepRecords.map((item) => item.result).filter(isResult));
      record.status = "failed";
      record.nextStep = failedIndex + 1;
      await saveWorkflowIfConfigured(options.workflowRoot, record);
      return resultFromRecord(record, "failed", rollback, failedIndex + 1);
    }
    index = batch[batch.length - 1] + 1;
  }

  record.status = "done";
  record.completedSteps = steps.length;
  record.nextStep = steps.length + 1;
  record.leaseExpiresAt = undefined;
  if (transaction && options.transactionRoot) {
    try {
      await finalizeOfficeTransaction(options.transactionRoot, transaction);
    } catch (error) {
      record.status = "failed";
      record.error = errorMessage(error);
      await saveWorkflowIfConfigured(options.workflowRoot, record);
      return resultFromRecord(record, "failed", [], steps.length);
    }
  }
  await saveWorkflowIfConfigured(options.workflowRoot, record);
  return resultFromRecord(record, "done", []);
}

export async function getOfficeWorkflow(root: string, id: string): Promise<OfficeWorkflowRecord> {
  return getOfficeWorkflowRecord(root, id);
}

export async function listOfficeWorkflows(root: string): Promise<OfficeWorkflowRecord[]> {
  return listOfficeWorkflowRecords(root);
}

function validateWorkflow(steps: OfficeWorkflowStepInput[]): string | undefined {
  if (steps.length === 0) return "Office 工作流至少需要一个步骤";
  if (steps.length > 20) return "Office 工作流最多支持 20 个步骤";
  for (let index = 0; index < steps.length; index++) {
    const step = steps[index];
    if (!step.filePath) return `工作流第 ${index + 1} 步缺少 filePath`;
    if (["restoreBackup", "listBackups"].includes(step.operation)) return `工作流步骤不能直接调用 ${step.operation}`;
    if (step.id && steps.some((candidate, candidateIndex) => candidateIndex !== index && candidate.id === step.id)) {
      return `工作流步骤 id 重复: ${step.id}`;
    }
    if (step.parallelGroup && !/^[A-Za-z0-9_-]{1,40}$/.test(step.parallelGroup)) {
      return `工作流第 ${index + 1} 步 parallelGroup 无效`;
    }
  }
  for (let index = 0; index < steps.length;) {
    const batch = workflowBatchIndexes(steps, index);
    if (batch.length > 1) {
      const writeTargets = new Set<string>();
      for (const stepIndex of batch) {
        const step = steps[stepIndex];
        if (step.action === "inspect" || step.action === "validate") continue;
        const target = path.resolve(step.outputPath || step.filePath!).toLowerCase();
        if (writeTargets.has(target)) return `并行组 ${step.parallelGroup} 包含重复写入目标: ${target}`;
        writeTargets.add(target);
      }
    }
    index = batch[batch.length - 1] + 1;
  }
  return undefined;
}

async function rollbackTransaction(
  root: string,
  transaction: OfficeTransactionRecord,
  step: number,
  options: OfficeWorkflowRunOptions,
) {
  try {
    await undoOfficeTransaction(root, transaction.id, {
      prepareFiles: options.prepareTransaction,
      restoreFiles: options.restoreTransaction,
    });
    return [{ step, ok: true, summary: `已整体撤销事务 ${transaction.id}` }];
  } catch (error) {
    return [{ step, ok: false, summary: errorMessage(error) }];
  }
}

function collectStepArtifacts(result: OfficeActionResult): string[] {
  const artifacts = new Set<string>();
  if (result.outputPath && result.outputPath !== result.filePath) artifacts.add(result.outputPath);
  for (const change of result.changes) {
    if (change.target && path.isAbsolute(change.target) && path.extname(change.target)) artifacts.add(change.target);
  }
  return [...artifacts];
}

function workflowBatchIndexes(steps: OfficeWorkflowStepInput[], startIndex: number): number[] {
  const group = steps[startIndex].parallelGroup;
  if (!group) return [startIndex];
  const indexes: number[] = [];
  for (let index = startIndex; index < steps.length && steps[index].parallelGroup === group; index++) indexes.push(index);
  return indexes;
}

function updateWorkflowProgress(record: OfficeWorkflowRecord): void {
  record.completedSteps = record.stepRecords.filter((step) => step.status === "done" || step.status === "skipped").length;
  const next = record.stepRecords.findIndex((step) => step.status !== "done" && step.status !== "skipped");
  record.nextStep = next === -1 ? record.steps.length + 1 : next + 1;
}

function leaseExpiry(leaseMs?: number): string {
  const duration = Number.isFinite(leaseMs) ? Math.min(30 * 60_000, Math.max(30_000, Math.trunc(leaseMs!))) : 5 * 60_000;
  return new Date(Date.now() + duration).toISOString();
}

function resolutionFailure(step: OfficeWorkflowStepInput, message: string): OfficeActionResult {
  return {
    status: "failed",
    engine: step.preferEngine || "com",
    app: step.app,
    action: step.action,
    operation: step.operation,
    filePath: step.filePath,
    outputPath: step.outputPath,
    target: step.target,
    summary: `工作流步骤参数解析失败: ${message}`,
    changes: [],
    error: message,
  };
}

function resultFromRecord(
  record: OfficeWorkflowRecord,
  status: OfficeWorkflowResult["status"],
  rollback: OfficeWorkflowResult["rollback"],
  failedStep?: number,
): OfficeWorkflowResult {
  return {
    status,
    workflowId: record.id,
    transactionId: record.transactionId,
    completedSteps: record.completedSteps,
    nextStep: record.nextStep <= record.steps.length ? record.nextStep : undefined,
    failedStep,
    results: record.stepRecords.map((item) => item.result).filter(isResult),
    stepRecords: record.stepRecords,
    rollback,
    error: record.error,
  };
}

function failedResult(error: string, record?: OfficeWorkflowRecord): OfficeWorkflowResult {
  return {
    status: "failed",
    workflowId: record?.id,
    transactionId: record?.transactionId,
    completedSteps: record?.completedSteps || 0,
    nextStep: record?.nextStep,
    results: record?.stepRecords.map((item) => item.result).filter(isResult) || [],
    stepRecords: record?.stepRecords || [],
    rollback: [],
    error,
  };
}

function isResult(value: OfficeActionResult | undefined): value is OfficeActionResult {
  return Boolean(value);
}

async function saveWorkflowIfConfigured(root: string | undefined, record: OfficeWorkflowRecord): Promise<void> {
  if (root) await saveOfficeWorkflowRecord(root, record);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
