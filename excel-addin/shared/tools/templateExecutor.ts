import type { HostAdapter } from "../host/types";
import type { WorkbookTemplateApplyInput, WorkbookTemplatePreset } from "../host/workbookTemplateTypes";
import {
  WORKBOOK_TEMPLATE_DEFAULT_FONT_NAME,
  WORKBOOK_TEMPLATE_DEFAULT_FONT_SIZE,
  WORKBOOK_TEMPLATE_PRESETS,
} from "../host/workbookTemplateTypes";
import type { ToolCall, ToolResult } from "./types";

function rejectUnknown(args: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function fromHost(
  tool: ToolCall["name"],
  result: { ok: boolean; data?: unknown; reason?: string; unsupported?: boolean },
): ToolResult {
  if (result.ok) return { ok: true, tool, data: result.data };
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      error: result.reason ?? "host failed",
      detail: result,
      unsupported: true,
    };
  }
  return { ok: false, tool, error: result.reason ?? "host failed", detail: result };
}

function optionalBoolean(args: Record<string, unknown>, key: string, fallback: boolean): boolean {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return fallback;
  }
  if (typeof args[key] !== "boolean") throw new Error(`${key} must be a boolean`);
  return args[key] as boolean;
}

function parsePreset(args: Record<string, unknown>): WorkbookTemplatePreset {
  if (!Object.prototype.hasOwnProperty.call(args, "preset") || args.preset === undefined) {
    return "professional";
  }
  const value = args.preset;
  if (typeof value !== "string" || !(WORKBOOK_TEMPLATE_PRESETS as readonly string[]).includes(value)) {
    throw new Error("preset must be professional|financial|dashboard|minimal");
  }
  return value as WorkbookTemplatePreset;
}

function parseFontName(args: Record<string, unknown>): string {
  if (!Object.prototype.hasOwnProperty.call(args, "fontName") || args.fontName === undefined) {
    return WORKBOOK_TEMPLATE_DEFAULT_FONT_NAME;
  }
  const value = args.fontName;
  if (typeof value !== "string") throw new Error("fontName must be a string");
  if (value.trim() === "" || value.length > 255) {
    throw new Error("fontName must be non-empty and ≤255");
  }
  return value;
}

function parseFontSize(args: Record<string, unknown>): number {
  if (!Object.prototype.hasOwnProperty.call(args, "fontSize") || args.fontSize === undefined) {
    return WORKBOOK_TEMPLATE_DEFAULT_FONT_SIZE;
  }
  const value = args.fontSize;
  if (typeof value !== "number" || !Number.isFinite(value) || value < 1 || value > 409) {
    throw new Error("fontSize must be a finite number in 1..409");
  }
  return value;
}

function parseFreezeRows(args: Record<string, unknown>): number {
  if (!Object.prototype.hasOwnProperty.call(args, "freezeRows") || args.freezeRows === undefined) {
    return 1;
  }
  const value = args.freezeRows;
  if (typeof value !== "number" || !Number.isInteger(value) || value < 0 || value > 1_048_576) {
    throw new Error("freezeRows must be an integer in 0..1048576");
  }
  return value;
}

function parseSheetNames(args: Record<string, unknown>): string[] | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "sheetNames") || args.sheetNames === undefined) {
    return undefined;
  }
  const value = args.sheetNames;
  if (!Array.isArray(value)) throw new Error("sheetNames must be an array");
  if (value.length > 500) throw new Error("sheetNames exceeds max 500");
  const out: string[] = [];
  const seen = new Set<string>();
  for (const item of value) {
    if (typeof item !== "string") throw new Error("sheetNames items must be strings");
    if (item.trim() === "") throw new Error("sheetNames items must be non-empty");
    if (item.length > 255) throw new Error("sheetNames item exceeds max length 255");
    const key = item.toLowerCase();
    if (seen.has(key)) throw new Error(`duplicate sheetNames entry: ${item}`);
    seen.add(key);
    out.push(item);
  }
  return out;
}

function parseApplyInput(args: Record<string, unknown>): WorkbookTemplateApplyInput {
  rejectUnknown(args, [
    "preset",
    "sheetNames",
    "allSheets",
    "fontName",
    "fontSize",
    "autoFit",
    "showGridlines",
    "freezeRows",
  ]);
  // Reject explicit nulls on known keys
  for (const key of Object.keys(args)) {
    if (args[key] === null) throw new Error(`${key} must not be null`);
  }
  return {
    preset: parsePreset(args),
    sheetNames: parseSheetNames(args),
    allSheets: optionalBoolean(args, "allSheets", true),
    fontName: parseFontName(args),
    fontSize: parseFontSize(args),
    autoFit: optionalBoolean(args, "autoFit", true),
    showGridlines: optionalBoolean(args, "showGridlines", false),
    freezeRows: parseFreezeRows(args),
  };
}

export async function executeTemplateTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "workbook.template.apply") {
    const input = parseApplyInput(call.arguments);
    return fromHost(call.name, await host.applyWorkbookTemplate(input));
  }
  if (call.name === "workbook.template.capture") {
    rejectUnknown(call.arguments, []);
    for (const key of Object.keys(call.arguments)) {
      if (call.arguments[key] === null) throw new Error(`${key} must not be null`);
    }
    return fromHost(call.name, await host.captureWorkbookTemplate());
  }
  return null;
}
