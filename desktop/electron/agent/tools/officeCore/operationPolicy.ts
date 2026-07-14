import type { OfficeActionKind } from "./types";

const SAFE_ACTION_OPERATIONS = new Set([
  "inspectFile",
  "layout",
  "tables",
  "listBackups",
  "traceFormulaDependencies",
  "inspectFormulaDependencies",
  "inspectFormulaBackups",
  "inspectFormulaProtection",
  "inspectPrintSettings",
  "inspectDocumentFormatting",
  "inspectReferences",
  "inspectRevisions",
  "inspectContentControls",
  "inspectPowerQueries",
  "inspectCharts",
  "inspectWorkbookObjects",
  "captureWorkbookTemplate",
  "inspectWorkbookFormatting",
  "inspectPresentationTheme",
  "inspectSlideElements",
  "inspectAnimations",
  "inspectSpeakerNotes",
  "inspectLinkedOfficeContent",
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
