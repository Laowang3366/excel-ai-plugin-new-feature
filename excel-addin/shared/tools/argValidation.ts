import type { ToolName } from "./types";

/** Reject unknown object keys; default message matches existing tool executors. */
export function rejectUnknownFields(
  args: Record<string, unknown>,
  allowed: readonly string[],
  errorPrefix = "unknown field",
): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) {
      throw new Error(`${errorPrefix}: ${key}`);
    }
  }
}

/** Nested format object keys for range.format.write (schema closed set). */
export const RANGE_FORMAT_FIELD_KEYS = [
  "fontName",
  "fontSize",
  "fontBold",
  "fontColor",
  "fillColor",
  "numberFormat",
  "horizontalAlignment",
  "verticalAlignment",
  "wrapText",
] as const;

/**
 * Core/legacy tools whose model schema is strict but runtime previously ignored unknowns.
 * Keys and allowlists must match TOOL_DEFINITIONS parameters.properties exactly.
 */
export const CORE_TOOL_ARGUMENT_ALLOWLIST = {
  "host.status": [],
  "selection.get": [],
  "range.read": ["sheetName", "range", "expand"],
  "range.write": ["sheetName", "range", "values", "verify"],
  "range.clear": ["sheetName", "range"],
  "range.format.read": ["sheetName", "range"],
  "range.format.write": ["sheetName", "range", "format"],
  "formula.read": ["sheetName", "range"],
  "formula.write": ["sheetName", "range", "formula", "verify"],
  "formula.context": ["sheetName", "range"],
  "sheet.list": [],
  "sheet.operation": ["operation", "sheetName", "newName", "position"],
  "sheet.add": ["sheetName"],
  "sheet.rename": ["sheetName", "newName"],
  "sheet.delete": ["sheetName"],
  "table.list": ["sheetName"],
  "table.create": ["sheetName", "range", "name", "hasHeaders"],
  "table.delete": ["sheetName", "tableName"],
  "workbook.inspect": [],
  "workbook.objects.inspect": ["maxItemsPerCategory", "sheetName"],
  "workbook.save": [],
  "conditionalFormat.list": ["sheetName", "range"],
  "conditionalFormat.add": ["sheetName", "range", "rule"],
  "conditionalFormat.delete": ["sheetName", "range", "id"],
  "dataValidation.read": ["sheetName", "range"],
  "dataValidation.write": ["sheetName", "range", "rule"],
  "dataValidation.clear": ["sheetName", "range"],
} as const satisfies Record<string, readonly string[]>;

export type CoreUnknownArgToolName = keyof typeof CORE_TOOL_ARGUMENT_ALLOWLIST;

export function isCoreUnknownArgToolName(name: string): name is CoreUnknownArgToolName {
  return Object.prototype.hasOwnProperty.call(CORE_TOOL_ARGUMENT_ALLOWLIST, name);
}

/** Top-level unknown rejection for the 25 core legacy paths. */
export function rejectUnknownCoreToolArguments(
  name: ToolName,
  args: Record<string, unknown>,
): void {
  if (!isCoreUnknownArgToolName(name)) return;
  rejectUnknownFields(args, CORE_TOOL_ARGUMENT_ALLOWLIST[name]);
}

/**
 * Nested format unknown rejection for range.format.write.
 * Caller must already ensure format is a plain object when present.
 */
export function rejectUnknownRangeFormatFields(format: Record<string, unknown>): void {
  rejectUnknownFields(format, RANGE_FORMAT_FIELD_KEYS, "unknown format field");
}

/** Ident: name/A1/enum — trim before host; empty after trim fails. */
export function requireIdent(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

/**
 * Optional Ident: missing/null/"" → undefined; non-string → Invalid;
 * pure whitespace → error (not silent omit); else trimmed.
 */
export function optionalIdent(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = value.trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

/**
 * Value content (formula/refersTo): emptiness via trim; return original string.
 */
export function requireValueString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value;
}

/**
 * Optional value: missing/null/"" → undefined; non-string → Invalid;
 * pure whitespace → error; else original string (spaces preserved).
 */
export function optionalValueString(
  args: Record<string, unknown>,
  key: string,
): string | undefined {
  const value = args[key];
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new Error(`Invalid string argument: ${key}`);
  if (value.trim() === "") throw new Error(`${key} must be non-empty`);
  return value;
}
