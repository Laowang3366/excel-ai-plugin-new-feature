import type { OfficeActionBridge } from "../contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "./types";

export interface OfficeWorkflowCondition {
  step: number | string;
  status?: "done" | "failed" | "skipped";
  dataPath?: string;
  equals?: unknown;
  exists?: boolean;
}

export interface OfficeWorkflowRetry {
  maxAttempts?: number;
  delayMs?: number;
}

export interface OfficeWorkflowStepInput extends OfficeActionInput {
  id?: string;
  when?: OfficeWorkflowCondition;
  retry?: OfficeWorkflowRetry;
  timeoutMs?: number;
  parallelGroup?: string;
}

export interface WorkflowResultReference {
  step: number;
  id?: string;
  status: "pending" | "running" | "done" | "failed" | "skipped";
  result?: OfficeActionResult;
}

export function shouldRunWorkflowStep(
  condition: OfficeWorkflowCondition | undefined,
  records: WorkflowResultReference[],
): boolean {
  if (!condition) return true;
  const record = findRecord(condition.step, records);
  if (!record) return false;
  if (condition.status && record.status !== condition.status) return false;
  if (!condition.dataPath) return true;
  const value = readPath(record.result?.data, condition.dataPath);
  if (condition.exists !== undefined && (value !== undefined) !== condition.exists) return false;
  return condition.equals === undefined || deepEqual(value, condition.equals);
}

export function resolveWorkflowStep(
  step: OfficeWorkflowStepInput,
  records: WorkflowResultReference[],
  variables: Record<string, unknown> = {},
): OfficeActionInput {
  const resolved = resolveValue(step, records, variables) as OfficeWorkflowStepInput;
  const params = { ...(resolved.params || {}) };
  if (resolved.timeoutMs !== undefined) params.actionTimeoutMs = resolved.timeoutMs;
  return {
    app: resolved.app,
    action: resolved.action,
    operation: resolved.operation,
    filePath: resolved.filePath,
    outputPath: resolved.outputPath,
    target: resolved.target,
    preferEngine: resolved.preferEngine,
    ...(Object.keys(params).length > 0 ? { params } : {}),
  };
}

export function resolveWorkflowVariables(
  steps: OfficeWorkflowStepInput[],
  variables: Record<string, unknown>,
): OfficeWorkflowStepInput[] {
  return resolveVariablesOnly(steps, variables) as OfficeWorkflowStepInput[];
}

export async function executeWorkflowStepWithRetry(
  bridge: OfficeActionBridge,
  step: OfficeActionInput,
  retry: OfficeWorkflowRetry | undefined,
): Promise<{ result: OfficeActionResult; attempts: number }> {
  const maxAttempts = clampInteger(retry?.maxAttempts, 1, 5, 1);
  const delayMs = clampInteger(retry?.delayMs, 0, 10_000, 300);
  let result = failedStepResult(step, "Office 操作未执行");
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    result = await executeOnce(bridge, step);
    if (result.status === "done" || result.status === "unsupported" || result.status === "needsCom") {
      return { result, attempts: attempt };
    }
    if (attempt < maxAttempts && delayMs > 0) await delay(delayMs * attempt);
  }
  return { result, attempts: maxAttempts };
}

function resolveValue(value: unknown, records: WorkflowResultReference[], variables: Record<string, unknown>): unknown {
  if (typeof value === "string") {
    const withVariables = replaceVariables(value, variables);
    return withVariables.replace(/\{\{steps\.([^.}]+)\.([^}]+)\}\}/g, (_match, selector: string, dataPath: string) => {
      const record = findRecord(/^\d+$/.test(selector) ? Number(selector) : selector, records);
      const resolved = readPath(record?.result, dataPath);
      if (resolved === undefined || resolved === null) throw new Error(`工作流占位符没有值: steps.${selector}.${dataPath}`);
      return typeof resolved === "string" ? resolved : JSON.stringify(resolved);
    });
  }
  if (Array.isArray(value)) return value.map((item) => resolveValue(item, records, variables));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolveValue(item, records, variables)]));
  }
  return value;
}

function resolveVariablesOnly(value: unknown, variables: Record<string, unknown>): unknown {
  if (typeof value === "string") return replaceVariables(value, variables);
  if (Array.isArray(value)) return value.map((item) => resolveVariablesOnly(item, variables));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, resolveVariablesOnly(item, variables)]));
  }
  return value;
}

function replaceVariables(value: string, variables: Record<string, unknown>): string {
  return value.replace(/\{\{vars\.([^}]+)\}\}/g, (_match, dataPath: string) => {
    const resolved = readPath(variables, dataPath);
    if (resolved === undefined || resolved === null) throw new Error(`工作流变量没有值: vars.${dataPath}`);
    return typeof resolved === "string" ? resolved : JSON.stringify(resolved);
  });
}

function findRecord(step: number | string, records: WorkflowResultReference[]): WorkflowResultReference | undefined {
  return typeof step === "number"
    ? records.find((record) => record.step === step)
    : records.find((record) => record.id === step);
}

function readPath(value: unknown, dataPath: string): unknown {
  return dataPath.split(".").filter(Boolean).reduce<unknown>((current, key) => {
    if (!current || typeof current !== "object") return undefined;
    if (Array.isArray(current) && /^\d+$/.test(key)) return current[Number(key)];
    return (current as Record<string, unknown>)[key];
  }, value);
}

async function executeOnce(bridge: OfficeActionBridge, step: OfficeActionInput): Promise<OfficeActionResult> {
  try {
    const result = await bridge.executeAction(step);
    return result || failedStepResult(step, "Office 操作未返回执行结果");
  } catch (error) {
    return failedStepResult(step, error instanceof Error ? error.message : String(error));
  }
}

function failedStepResult(step: OfficeActionInput, message: string): OfficeActionResult {
  return {
    status: "failed",
    engine: step.preferEngine || "com",
    app: step.app,
    action: step.action,
    operation: step.operation,
    filePath: step.filePath,
    outputPath: step.outputPath,
    target: step.target,
    summary: `Office 工作流步骤执行失败: ${message}`,
    changes: [],
    error: message,
  };
}

function clampInteger(value: number | undefined, min: number, max: number, fallback: number): number {
  return Number.isInteger(value) ? Math.min(max, Math.max(min, value!)) : fallback;
}

function deepEqual(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
