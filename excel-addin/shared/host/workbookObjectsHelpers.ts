import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import type { ObjectCategoryResult } from "./workbookObjectsTypes";

function cmpStr(a: string, b: string): number {
  return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
}

/** sheetName (as stored) -> workbook position/index; lower sorts first. */
export type SheetOrderIndex = ReadonlyMap<string, number>;

export function buildSheetOrder(sheets: readonly SheetInfo[]): SheetOrderIndex {
  const map = new Map<string, number>();
  for (const sheet of sheets) {
    map.set(sheet.name, sheet.index);
  }
  return map;
}

/** Resolve position; unknown sheets sort last (stable via name). */
export function sheetPosition(order: SheetOrderIndex, sheetName: string): number {
  const direct = order.get(sheetName);
  if (direct !== undefined) return direct;
  for (const [name, index] of order) {
    if (name.localeCompare(sheetName, undefined, { sensitivity: "accent" }) === 0) {
      return index;
    }
  }
  return Number.MAX_SAFE_INTEGER;
}

export function sortSheets(sheets: SheetInfo[]): SheetInfo[] {
  return [...sheets].sort((a, b) => {
    if (a.index !== b.index) return a.index - b.index;
    return cmpStr(a.name, b.name);
  });
}

export function sortTables(items: TableInfo[], order: SheetOrderIndex): TableInfo[] {
  return [...items].sort((a, b) => {
    const pa = sheetPosition(order, a.sheetName);
    const pb = sheetPosition(order, b.sheetName);
    if (pa !== pb) return pa - pb;
    const sn = cmpStr(a.sheetName, b.sheetName);
    if (sn !== 0) return sn;
    return cmpStr(a.name, b.name);
  });
}

export function sortCharts(items: ChartInfo[], order: SheetOrderIndex): ChartInfo[] {
  return [...items].sort((a, b) => {
    const pa = sheetPosition(order, a.sheetName);
    const pb = sheetPosition(order, b.sheetName);
    if (pa !== pb) return pa - pb;
    const sn = cmpStr(a.sheetName, b.sheetName);
    if (sn !== 0) return sn;
    return cmpStr(a.name, b.name);
  });
}

export function sortShapes(items: ShapeInfo[], order: SheetOrderIndex): ShapeInfo[] {
  return [...items].sort((a, b) => {
    const pa = sheetPosition(order, a.sheetName);
    const pb = sheetPosition(order, b.sheetName);
    if (pa !== pb) return pa - pb;
    const sn = cmpStr(a.sheetName, b.sheetName);
    if (sn !== 0) return sn;
    return cmpStr(a.name, b.name);
  });
}

export function sortNamedRanges(items: NamedRangeInfo[], order: SheetOrderIndex): NamedRangeInfo[] {
  return [...items].sort((a, b) => {
    const scopeOrder = (s: NamedRangeInfo["scope"]) => (s === "workbook" ? 0 : 1);
    const so = scopeOrder(a.scope) - scopeOrder(b.scope);
    if (so !== 0) return so;
    if (a.scope === "worksheet" || b.scope === "worksheet") {
      const pa = sheetPosition(order, a.sheetName ?? "");
      const pb = sheetPosition(order, b.sheetName ?? "");
      if (pa !== pb) return pa - pb;
      const sn = cmpStr(a.sheetName ?? "", b.sheetName ?? "");
      if (sn !== 0) return sn;
    }
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
  const oneLine = message.replace(/\s+/g, " ").trim();
  return oneLine.length > 400 ? `${oneLine.slice(0, 400)}…` : oneLine;
}
