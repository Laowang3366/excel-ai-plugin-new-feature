import type { OfficeActionKind } from "./types";

const SAFE_ACTION_OPERATIONS = new Set([
  "inspectFile",
  "layout",
  "tables",
]);

export function officeActionOperationError(
  action: OfficeActionKind,
  operation: string
): string | undefined {
  if (
    (action === "inspect" || action === "validate") &&
    !SAFE_ACTION_OPERATIONS.has(operation)
  ) {
    return `${action} 仅允许只读 Office 操作；修改文件请使用 office.action.apply`;
  }
  return undefined;
}
