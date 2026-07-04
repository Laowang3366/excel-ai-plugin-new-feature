/**
 * Office 编辑事件提取
 *
 * 关联模块：
 * - components/office/OfficePreviewPanel.tsx: 渲染右侧编辑详情。
 * - store/chatStore.ts: 提供当前会话 TurnItem 列表。
 */

import type { ToolResultItem, TurnItem } from "../electronApi";

export type OfficeEditOperation = "officeAction";
export type OfficeEditDocumentType = "word" | "presentation" | "spreadsheet";

export interface OfficeEditEvent {
  id: string;
  timestamp: number;
  toolName: string;
  engine: "openxml";
  operation: OfficeEditOperation;
  documentType: OfficeEditDocumentType;
  filePath: string;
  outputPath?: string;
  summary: string;
  detail: Record<string, unknown>;
}

const TRACKED_OFFICE_TOOLS = new Set([
  "office.action.inspect",
  "office.action.apply",
  "office.action.validate",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function actionDocumentType(app: unknown): OfficeEditDocumentType | null {
  if (app === "word") return "word";
  if (app === "presentation") return "presentation";
  if (app === "excel") return "spreadsheet";
  return null;
}

function toOfficeEditEvent(item: ToolResultItem): OfficeEditEvent | null {
  if (item.isError || !TRACKED_OFFICE_TOOLS.has(item.toolName)) return null;
  if (!isRecord(item.result)) return null;
  if (item.result.engine !== "openxml") return null;
  return toOfficeActionEvent(item, item.result);
}

function toOfficeActionEvent(item: ToolResultItem, result: Record<string, unknown>): OfficeEditEvent | null {
  const documentType = actionDocumentType(result.app);
  if (!documentType) return null;
  if (typeof result.filePath !== "string") return null;
  if (typeof result.operation !== "string") return null;
  if (
    result.status !== "done" &&
    result.status !== "unsupported" &&
    result.status !== "needsCom" &&
    result.status !== "failed"
  ) return null;
  const app = typeof result.app === "string" ? result.app : documentType;

  return {
    id: item.id,
    timestamp: item.timestamp,
    toolName: item.toolName,
    engine: "openxml",
    operation: "officeAction",
    documentType,
    filePath: result.filePath,
    outputPath: typeof result.outputPath === "string" ? result.outputPath : undefined,
    summary: `Office action ${app}/${result.operation}：${result.status}`,
    detail: {
      summary: typeof result.summary === "string" ? result.summary : "",
      changes: Array.isArray(result.changes) ? result.changes : [],
      validation: result.validation,
      error: result.error,
    },
  };
}

export function collectOfficeEditEvents(items: TurnItem[]): OfficeEditEvent[] {
  return items
    .filter((item): item is ToolResultItem => item.type === "tool_result")
    .map(toOfficeEditEvent)
    .filter((event): event is OfficeEditEvent => Boolean(event));
}

export type OfficePreviewToggleLocation = "chat-header" | "panel-header";

export function getOfficePreviewToggleLocation(isOpen: boolean): OfficePreviewToggleLocation {
  return isOpen ? "panel-header" : "chat-header";
}
