/**
 * Office action 统一协议类型
 *
 * 关联模块：
 * - locator.ts: 解析 action target。
 * - results.ts: 生成统一 action 结果。
 * - officeActionAdapter.ts: 按协议路由到 Open XML 或 COM。
 */

export type OfficeActionApp = "excel" | "word" | "presentation";
export type OfficeActionKind = "inspect" | "edit" | "style" | "insert" | "snapshot" | "validate";
export type OfficeActionStatus = "done" | "unsupported" | "needsCom" | "failed";
export type OfficeActionEngine = "openxml" | "com";

export interface OfficeLocator {
  kind: string;
  value: string;
  sheetName?: string;
  address?: string;
  index?: number;
}

export interface OfficeActionInput {
  app: OfficeActionApp;
  action: OfficeActionKind;
  operation: string;
  filePath?: string;
  outputPath?: string;
  target?: string;
  preferEngine?: OfficeActionEngine;
  params?: Record<string, unknown>;
}

export interface OfficeActionValidation {
  ok: boolean;
  checks: Array<{ name: string; ok: boolean; message: string }>;
}

export interface OfficeActionResult {
  status: OfficeActionStatus;
  engine: OfficeActionEngine;
  app: OfficeActionApp;
  action: OfficeActionKind;
  operation: string;
  filePath?: string;
  outputPath?: string;
  target?: string;
  summary: string;
  changes: Array<{ kind: string; target?: string; detail: string }>;
  validation?: OfficeActionValidation;
  error?: string;
  data?: unknown;
}
