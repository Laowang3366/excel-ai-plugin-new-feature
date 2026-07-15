import path from "node:path";

import type { OfficeActionResult } from "./types";
import { officeAdvancedOperationError } from "./operationPolicy";
import type { OfficeWorkflowStepInput } from "./workflowStepExecution";
import type { OfficeWorkflowRecord, OfficeWorkflowResult } from "./workflowTypes";

export function validateWorkflow(steps: OfficeWorkflowStepInput[]): string | undefined {
  if (steps.length === 0) return "Office 工作流至少需要一个步骤";
  if (steps.length > 20) return "Office 工作流最多支持 20 个步骤";
  for (let index = 0; index < steps.length; index++) {
    const step = steps[index];
    if (!step.filePath) return `工作流第 ${index + 1} 步缺少 filePath`;
    if (["restoreBackup", "listBackups"].includes(step.operation))
      return `工作流步骤不能直接调用 ${step.operation}`;
    const advancedOperationError = officeAdvancedOperationError(step);
    if (advancedOperationError) return `工作流第 ${index + 1} 步: ${advancedOperationError}`;
    if (
      step.id &&
      steps.some(
        (candidate, candidateIndex) => candidateIndex !== index && candidate.id === step.id,
      )
    ) {
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
        if (writeTargets.has(target))
          return `并行组 ${step.parallelGroup} 包含重复写入目标: ${target}`;
        writeTargets.add(target);
      }
    }
    index = batch[batch.length - 1] + 1;
  }
  return undefined;
}

export function collectStepArtifacts(result: OfficeActionResult): string[] {
  const artifacts = new Set<string>();
  if (result.outputPath && result.outputPath !== result.filePath) artifacts.add(result.outputPath);
  for (const change of result.changes) {
    if (change.target && path.isAbsolute(change.target) && path.extname(change.target))
      artifacts.add(change.target);
  }
  return [...artifacts];
}

export function workflowBatchIndexes(
  steps: OfficeWorkflowStepInput[],
  startIndex: number,
): number[] {
  const group = steps[startIndex].parallelGroup;
  if (!group) return [startIndex];
  const indexes: number[] = [];
  for (
    let index = startIndex;
    index < steps.length && steps[index].parallelGroup === group;
    index++
  )
    indexes.push(index);
  return indexes;
}

export function updateWorkflowProgress(record: OfficeWorkflowRecord): void {
  record.completedSteps = record.stepRecords.filter(
    (step) => step.status === "done" || step.status === "skipped",
  ).length;
  const next = record.stepRecords.findIndex(
    (step) => step.status !== "done" && step.status !== "skipped",
  );
  record.nextStep = next === -1 ? record.steps.length + 1 : next + 1;
}

export function leaseExpiry(leaseMs?: number): string {
  const duration = Number.isFinite(leaseMs)
    ? Math.min(30 * 60_000, Math.max(30_000, Math.trunc(leaseMs!)))
    : 5 * 60_000;
  return new Date(Date.now() + duration).toISOString();
}

export function resolutionFailure(
  step: OfficeWorkflowStepInput,
  message: string,
): OfficeActionResult {
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

export function resultFromRecord(
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

export function failedResult(error: string, record?: OfficeWorkflowRecord): OfficeWorkflowResult {
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

export function isResult(value: OfficeActionResult | undefined): value is OfficeActionResult {
  return Boolean(value);
}

export function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
