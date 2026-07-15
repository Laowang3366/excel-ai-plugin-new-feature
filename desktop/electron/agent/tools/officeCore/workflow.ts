import { randomUUID } from "node:crypto";

import type { OfficeActionBridge } from "../contracts/office";
import {
  beginOfficeTransaction,
  finalizeOfficeTransaction,
  getOfficeTransaction,
  listOfficeTransactionPaths,
  recordOfficeTransactionResult,
  undoOfficeTransaction,
} from "./transactionJournal";
import type { OfficeTransactionRecord } from "./transactionTypes";
import type { OfficeActionInput } from "./types";
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
import {
  collectStepArtifacts,
  errorMessage,
  failedResult,
  isResult,
  leaseExpiry,
  resolutionFailure,
  resultFromRecord,
  updateWorkflowProgress,
  validateWorkflow,
  workflowBatchIndexes,
} from "./workflowHelpers";
import type {
  OfficeWorkflowRecord,
  OfficeWorkflowResult,
  OfficeWorkflowRunOptions,
} from "./workflowTypes";

export { requestOfficeWorkflowCancellation } from "./workflowRecordStore";

export async function runOfficeWorkflow(
  bridge: OfficeActionBridge,
  requestedSteps: OfficeWorkflowStepInput[],
  options: OfficeWorkflowRunOptions = {},
): Promise<OfficeWorkflowResult> {
  if (!options.workflowRoot) return runOfficeWorkflowUnlocked(bridge, requestedSteps, options);
  const workflowId = options.workflowId || randomUUID();
  let lock: Awaited<ReturnType<typeof acquireOfficeWorkflowLock>>;
  try {
    lock = await acquireOfficeWorkflowLock(options.workflowRoot, workflowId);
  } catch (error) {
    return failedResult(errorMessage(error));
  }
  const stopHeartbeat = startOfficeWorkflowLockHeartbeat(lock);
  try {
    await lock.assertOwned();
    return await runOfficeWorkflowUnlocked(
      bridge,
      requestedSteps,
      { ...options, workflowId },
      lock.assertOwned,
    );
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
    if (!options.workflowRoot || !options.workflowId)
      return failedResult("继续工作流需要 workflowRoot 和 workflowId");
    try {
      record = await getOfficeWorkflowRecord(options.workflowRoot, options.workflowId);
    } catch (error) {
      return failedResult(errorMessage(error));
    }
    const leaseExpired = !record.leaseExpiresAt || Date.parse(record.leaseExpiresAt) <= Date.now();
    const recoverRunning =
      record.status === "running" && (options.recoverRunning === true || leaseExpired);
    if (!["paused", "failed", "cancelled"].includes(record.status) && !recoverRunning) {
      return failedResult("只有暂停、失败、已取消或租约过期的工作流可以继续", record);
    }
    if (recoverRunning) {
      for (const stepRecord of record.stepRecords) {
        if (stepRecord.status === "running") stepRecord.status = "failed";
      }
      record.status = "paused";
    }
    if (
      requestedSteps.length > 0 &&
      JSON.stringify(requestedSteps) !== JSON.stringify(record.sourceSteps || record.steps)
    ) {
      return failedResult("继续工作流时提供的 steps 与原执行计划不一致", record);
    }
    steps = record.steps;
    if (options.transactionRoot && record.transactionId) {
      try {
        transaction = await getOfficeTransaction(options.transactionRoot, record.transactionId);
      } catch (error) {
        return failedResult(errorMessage(error), record);
      }
    }
  }

  if (!resumed && options.variables) {
    try {
      steps = resolveWorkflowVariables(steps, options.variables);
    } catch (error) {
      return failedResult(errorMessage(error));
    }
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
      stepRecords: steps.map((step, index) => ({
        step: index + 1,
        id: step.id,
        status: "pending",
        artifacts: [],
      })),
      completedSteps: 0,
      nextStep: 1,
      variables: options.variables,
    };
    if (options.transactionRoot) {
      if (options.prepareTransaction) {
        try {
          await options.prepareTransaction(listOfficeTransactionPaths(steps));
        } catch (error) {
          return failedResult(`准备已打开的 Office 文档失败: ${errorMessage(error)}`, record);
        }
      }
      transaction = await beginOfficeTransaction({
        root: options.transactionRoot,
        steps,
        workflowId: id,
      });
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
        const rollback =
          transaction && options.transactionRoot
            ? await rollbackTransaction(options.transactionRoot, transaction, index + 1, options)
            : await rollbackOfficeResults(
                bridge,
                steps,
                record.stepRecords.map((item) => item.result).filter(isResult),
              );
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
        runnable.push({
          index: stepIndex,
          resolved: { ...stepRecord.resolvedStep, transactionContext: "workflow" },
        });
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

    const executions = await Promise.all(
      runnable.map(async (item) => ({
        ...item,
        execution: await executeWorkflowStepWithRetry(
          bridge,
          item.resolved,
          steps[item.index].retry,
        ),
      })),
    );
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
    const failedIndex = batch.find(
      (stepIndex) => record.stepRecords[stepIndex].status === "failed",
    );
    if (failedIndex !== undefined) {
      const failedRecord = record.stepRecords[failedIndex];
      record.error =
        failedRecord.result?.error || failedRecord.result?.summary || "Office 工作流步骤失败";
      const failureMode = options.failureMode || (options.workflowRoot ? "pause" : "rollback");
      if (failureMode === "pause") {
        record.status = "paused";
        record.nextStep = failedIndex + 1;
        await saveWorkflowIfConfigured(options.workflowRoot, record);
        return resultFromRecord(record, "paused", [], failedIndex + 1);
      }
      const rollback =
        transaction && options.transactionRoot
          ? await rollbackTransaction(
              options.transactionRoot,
              transaction,
              failedIndex + 1,
              options,
            )
          : await rollbackOfficeResults(
              bridge,
              steps,
              record.stepRecords.map((item) => item.result).filter(isResult),
            );
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

async function saveWorkflowIfConfigured(
  root: string | undefined,
  record: OfficeWorkflowRecord,
): Promise<void> {
  if (root) await saveOfficeWorkflowRecord(root, record);
}
