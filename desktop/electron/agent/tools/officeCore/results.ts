/**
 * Office action 结果构造工具
 *
 * 关联模块：
 * - types.ts: 统一结果类型。
 * - officeActionAdapter.ts: 使用这些 helper 保持状态语义一致。
 */

import type { OfficeActionInput, OfficeActionResult } from "./types";

type ResultBase = Pick<OfficeActionResult, "engine" | "app" | "action" | "operation" | "summary"> &
  Partial<Pick<OfficeActionResult, "filePath" | "outputPath" | "target" | "changes" | "validation" | "error" | "data">>;

export function doneResult(input: ResultBase): OfficeActionResult {
  return { status: "done", changes: [], ...input };
}

export function unsupportedResult(input: Omit<ResultBase, "engine">): OfficeActionResult {
  return { status: "unsupported", engine: "openxml", changes: [], ...input };
}

export function needsComResult(input: Omit<ResultBase, "engine">): OfficeActionResult {
  return { status: "needsCom", engine: "openxml", changes: [], ...input };
}

export function failedResult(action: OfficeActionInput, error: unknown): OfficeActionResult {
  return {
    status: "failed",
    engine: action.preferEngine || "openxml",
    app: action.app,
    action: action.action,
    operation: action.operation,
    filePath: action.filePath,
    outputPath: action.outputPath,
    target: action.target,
    summary: "Office action 执行失败",
    changes: [],
    error: error instanceof Error ? error.message : String(error),
  };
}
