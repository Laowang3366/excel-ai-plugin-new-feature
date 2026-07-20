import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import type { ObjectCategoryResult } from "./workbookObjectsTypes";

function cmpStr(a: string, b: string): number {
  return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
}

export function sortSheets(sheets: SheetInfo[]): SheetInfo[] {
  return [...sheets].sort((a, b) => {
    if (a.index !== b.index) return a.index - b.index;
    return cmpStr(a.name, b.name);
  });
}

export function sortTables(items: TableInfo[]): TableInfo[] {
  return [...items].sort((a, b) => {
    const s = cmpStr(a.sheetName, b.sheetName);
    if (s !== 0) return s;
    return cmpStr(a.name, b.name);
  });
}

export function sortCharts(items: ChartInfo[]): ChartInfo[] {
  return [...items].sort((a, b) => {
    const s = cmpStr(a.sheetName, b.sheetName);
    if (s !== 0) return s;
    return cmpStr(a.name, b.name);
  });
}

export function sortNamedRanges(items: NamedRangeInfo[]): NamedRangeInfo[] {
  return [...items].sort((a, b) => {
    const scopeOrder = (s: NamedRangeInfo["scope"]) => (s === "workbook" ? 0 : 1);
    const so = scopeOrder(a.scope) - scopeOrder(b.scope);
    if (so !== 0) return so;
    const sheetA = a.sheetName ?? "";
    const sheetB = b.sheetName ?? "";
    const ss = cmpStr(sheetA, sheetB);
    if (ss !== 0) return ss;
    return cmpStr(a.name, b.name);
  });
}

export function sortShapes(items: ShapeInfo[]): ShapeInfo[] {
  return [...items].sort((a, b) => {
    const s = cmpStr(a.sheetName, b.sheetName);
    if (s !== 0) return s;
    return cmpStr(a.name, b.name);
  });
}

/** Apply stable sort then truncate; totalCount is always the full sorted length. */
export function availableCategory<T>(
  items: T[],
  maxItems: number,
  sort: (items: T[]) => T[],
  extra?: { limitations?: string[] },
): ObjectCategoryResult<T> {
  const sorted = sort(items);
  const totalCount = sorted.length;
  const truncated = totalCount > maxItems;
  return {
    status: "available",
    totalCount,
    items: truncated ? sorted.slice(0, maxItems) : sorted,
    truncated,
    limitations: extra?.limitations,
  };
}

export function unsupportedCategory<T>(
  reason: string,
  evidence?: string,
  limitations?: string[],
): ObjectCategoryResult<T> {
  return {
    status: "unsupported",
    totalCount: null,
    items: [],
    truncated: false,
    reason,
    evidence,
    limitations,
  };
}

export function failedCategory<T>(
  reason: string,
  evidence?: string,
  limitations?: string[],
): ObjectCategoryResult<T> {
  return {
    status: "failed",
    totalCount: null,
    items: [],
    truncated: false,
    reason,
    evidence,
    limitations,
  };
}

/** Map a host list result into a category without treating unsupported as empty success. */
export function categoryFromHostList<T>(
  result: HostResult<T[]>,
  maxItems: number,
  sort: (items: T[]) => T[],
): ObjectCategoryResult<T> {
  if (result.ok) {
    return availableCategory(result.data, maxItems, sort);
  }
  if (result.unsupported) {
    return unsupportedCategory(result.reason, result.evidence);
  }
  return failedCategory(result.reason, result.evidence);
}

export function sanitizeInventoryMessage(message: string): string {
  // Keep messages short and free of accidental path dumps.
  const oneLine = message.replace(/\s+/g, " ").trim();
  return oneLine.length > 400 ? `${oneLine.slice(0, 400)}…` : oneLine;
}
