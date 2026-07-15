import { officeAdvancedOperationError } from "../officeCore/operationPolicy";
import type {
  OfficeWorkflowCondition,
  OfficeWorkflowStepInput,
} from "../officeCore/workflowStepExecution";
import type { OfficeActionApp, OfficeActionEngine, OfficeActionKind } from "../officeCore/types";

export function parseWorkflowSteps(value: unknown): OfficeWorkflowStepInput[] | string {
  if (!Array.isArray(value)) return "参数 steps 必须是数组";
  const steps: OfficeWorkflowStepInput[] = [];
  for (let index = 0; index < value.length; index++) {
    const raw = value[index];
    if (!raw || typeof raw !== "object" || Array.isArray(raw))
      return `工作流第 ${index + 1} 步必须是对象`;
    const item = raw as Record<string, unknown>;
    if (!isOfficeActionApp(item.app)) return `工作流第 ${index + 1} 步 app 无效`;
    if (!isOfficeActionKind(item.action)) return `工作流第 ${index + 1} 步 action 无效`;
    if (typeof item.operation !== "string" || typeof item.filePath !== "string")
      return `工作流第 ${index + 1} 步需要 operation 和 filePath`;
    const step: OfficeWorkflowStepInput = {
      app: item.app,
      action: item.action,
      operation: item.operation,
      filePath: item.filePath,
    };
    if (typeof item.outputPath === "string") step.outputPath = item.outputPath;
    if (typeof item.target === "string") step.target = item.target;
    if (isOfficeActionEngine(item.preferEngine)) step.preferEngine = item.preferEngine;
    if (item.params && typeof item.params === "object" && !Array.isArray(item.params))
      step.params = item.params as Record<string, unknown>;
    const advancedOperationError = officeAdvancedOperationError(step);
    if (advancedOperationError) return `工作流第 ${index + 1} 步: ${advancedOperationError}`;
    if (typeof item.id === "string") step.id = item.id;
    if (typeof item.parallelGroup === "string") step.parallelGroup = item.parallelGroup;
    if (typeof item.timeoutMs === "number") step.timeoutMs = item.timeoutMs;
    if (item.retry !== undefined) {
      if (!item.retry || typeof item.retry !== "object" || Array.isArray(item.retry))
        return `工作流第 ${index + 1} 步 retry 必须是对象`;
      const retry = item.retry as Record<string, unknown>;
      step.retry = {
        maxAttempts: typeof retry.maxAttempts === "number" ? retry.maxAttempts : undefined,
        delayMs: typeof retry.delayMs === "number" ? retry.delayMs : undefined,
      };
    }
    if (item.when !== undefined) {
      const condition = parseWorkflowCondition(item.when);
      if (typeof condition === "string") return `工作流第 ${index + 1} 步 ${condition}`;
      step.when = condition;
    }
    steps.push(step);
  }
  return steps;
}

export function parseWorkflowCondition(value: unknown): OfficeWorkflowCondition | string {
  if (!value || typeof value !== "object" || Array.isArray(value)) return "when 必须是对象";
  const item = value as Record<string, unknown>;
  if (typeof item.step !== "number" && typeof item.step !== "string")
    return "when.step 必须是步骤序号或 ID";
  if (item.status !== undefined && !["done", "failed", "skipped"].includes(String(item.status)))
    return "when.status 无效";
  return {
    step: item.step,
    status: item.status as OfficeWorkflowCondition["status"],
    dataPath: typeof item.dataPath === "string" ? item.dataPath : undefined,
    ...(Object.prototype.hasOwnProperty.call(item, "equals") ? { equals: item.equals } : {}),
    exists: typeof item.exists === "boolean" ? item.exists : undefined,
  };
}

export function isOfficeActionApp(value: unknown): value is OfficeActionApp {
  return value === "excel" || value === "word" || value === "presentation";
}

function isOfficeActionKind(value: unknown): value is OfficeActionKind {
  return (
    value === "inspect" ||
    value === "edit" ||
    value === "style" ||
    value === "insert" ||
    value === "snapshot" ||
    value === "validate"
  );
}

function isOfficeActionEngine(value: unknown): value is OfficeActionEngine {
  return value === "openxml" || value === "com";
}
