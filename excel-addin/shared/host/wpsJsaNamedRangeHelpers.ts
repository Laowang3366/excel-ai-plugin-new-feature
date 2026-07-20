import { getSheet, type WpsName, type WpsNames, type WpsWorkbook } from "./wpsJsaRuntime";
import type { HostResult, NamedRangeInfo, NamedRangeScope } from "./types";
import { fail, ok, unsupported } from "./types";

/** Assumed Names surface; not in bridge contract; not device-verified. */
export const NAMED_RANGE_EVIDENCE =
  "Assumed Names.Count/Item/Add + Name/RefersTo/Visible/Delete (desktop COM parity; not in bridge contract; not device-verified)";

export function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function sameName(a: string, b: string): boolean {
  return a.localeCompare(b, undefined, { sensitivity: "accent" }) === 0;
}

export function normalizeRefersTo(value: string): string {
  const trimmed = value.trim();
  if (trimmed === "") return trimmed;
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

export function readNamedRangeInfo(
  item: WpsName,
  scope: NamedRangeScope,
  sheetName?: string,
): NamedRangeInfo {
  return {
    name: String(item.Name ?? ""),
    refersTo: String(item.RefersTo ?? ""),
    scope,
    sheetName: scope === "worksheet" ? sheetName : undefined,
    visible: typeof item.Visible === "boolean" ? item.Visible : undefined,
  };
}

export function requireNamesCollection(
  capability: string,
  workbook: WpsWorkbook,
  scope: NamedRangeScope,
  sheetName?: string,
): HostResult<WpsNames> {
  if (scope === "worksheet") {
    if (!sheetName || sheetName.trim() === "") {
      return fail(
        capability,
        "wps-jsa",
        "sheetName is required for worksheet scope",
        NAMED_RANGE_EVIDENCE,
      );
    }
    const sheet = getSheet(workbook, sheetName);
    if (!sheet) {
      return unsupported(
        capability,
        "wps-jsa",
        `Sheet "${sheetName}" not found`,
        NAMED_RANGE_EVIDENCE,
      );
    }
    if (!sheet.Names) {
      return unsupported(
        capability,
        "wps-jsa",
        "Worksheet.Names is unavailable",
        NAMED_RANGE_EVIDENCE,
      );
    }
    const names = sheet.Names;
    if (typeof names.Count !== "number" || typeof names.Item !== "function") {
      return unsupported(
        capability,
        "wps-jsa",
        "Names.Count/Item is unavailable",
        NAMED_RANGE_EVIDENCE,
      );
    }
    return ok(names);
  }
  if (!workbook.Names) {
    return unsupported(
      capability,
      "wps-jsa",
      "Workbook.Names is unavailable",
      NAMED_RANGE_EVIDENCE,
    );
  }
  const names = workbook.Names;
  if (typeof names.Count !== "number" || typeof names.Item !== "function") {
    return unsupported(
      capability,
      "wps-jsa",
      "Names.Count/Item is unavailable",
      NAMED_RANGE_EVIDENCE,
    );
  }
  return ok(names);
}

export function findByName(names: WpsNames, name: string): WpsName | null {
  try {
    if (typeof names.Item === "function") {
      const item = names.Item(name);
      if (item && item.Name != null) return item;
    }
  } catch {
    // fall through to scan
  }
  const count = names.Count ?? 0;
  for (let i = 1; i <= count; i += 1) {
    try {
      const item = names.Item?.(i);
      if (item?.Name != null && sameName(String(item.Name), name)) return item;
    } catch {
      // continue
    }
  }
  return null;
}

export function hasName(names: WpsNames, name: string): boolean {
  return findByName(names, name) != null;
}
