/**
 * Excel/WPS 连接元数据工具。
 *
 * 关联模块：
 * - excelComBridge.ts: 连接状态和工作簿检查需要统一识别宿主。
 * - excelComBridge.test.ts: 锁定双宿主选择和 WPS 元数据归一化行为。
 */

export type SpreadsheetHost = "excel" | "wps";

export function resolveSpreadsheetHost(
  availableHosts: SpreadsheetHost[],
  selectedHost: SpreadsheetHost | null,
  currentHost: SpreadsheetHost | "unknown" = "unknown"
): SpreadsheetHost | null {
  if (selectedHost && availableHosts.includes(selectedHost)) {
    return selectedHost;
  }
  if (currentHost !== "unknown" && availableHosts.includes(currentHost)) {
    return currentHost;
  }
  return availableHosts.length === 1 ? availableHosts[0] : null;
}

export function normalizeWorkbookInspectMetadata(
  raw: unknown,
  host: SpreadsheetHost,
  version?: string
): unknown {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return raw;
  }

  return {
    ...(raw as Record<string, unknown>),
    host,
    name: host === "wps" ? "WPS 表格" : "Microsoft Excel",
    version: version || (raw as Record<string, unknown>).version || "unknown",
  };
}
