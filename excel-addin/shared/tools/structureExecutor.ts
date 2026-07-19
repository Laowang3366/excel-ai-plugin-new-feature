import type { HostAdapter, NamedRangeScope, SheetVisibility } from "../host/types";
import type { ToolCall, ToolResult } from "./types";
import {
  optionalValueString,
  requireIdent,
  requireValueString,
} from "./argValidation";

/** Password is special: null/"" omit; spaces and padding preserved raw. */
function optionalPassword(args: Record<string, unknown>): string | undefined {
  const value = args.password;
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new Error("Invalid string argument: password");
  return value;
}

/** Present key with only whitespace is an error (not silent omit). */
function optionalTrimmedName(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] == null) return undefined;
  if (typeof args[key] !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = (args[key] as string).trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

function resolveNamedScopeArgs(args: Record<string, unknown>): {
  scope: NamedRangeScope;
  sheetName?: string;
} {
  const scope = requireScope(args);
  const hasSheet =
    Object.prototype.hasOwnProperty.call(args, "sheetName") &&
    args.sheetName != null &&
    args.sheetName !== "";
  if (scope === "workbook") {
    if (hasSheet) throw new Error("sheetName is not allowed for workbook scope");
    return { scope };
  }
  const sheetName = optionalTrimmedName(args, "sheetName");
  if (!sheetName) throw new Error("sheetName is required for worksheet scope");
  return { scope, sheetName };
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  const value = args[key];
  if (value == null) return undefined;
  if (typeof value !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return value;
}

function requireVisibility(args: Record<string, unknown>): SheetVisibility {
  const value = requireIdent(args, "visibility");
  if (value !== "visible" && value !== "hidden" && value !== "veryHidden") {
    throw new Error("visibility must be visible|hidden|veryHidden");
  }
  return value;
}

function requireScope(args: Record<string, unknown>): NamedRangeScope {
  const value = requireIdent(args, "scope");
  if (value !== "workbook" && value !== "worksheet") {
    throw new Error("scope must be workbook|worksheet");
  }
  return value;
}

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

export async function executeStructureTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  switch (call.name) {
    case "sheet.visibility.get":
      rejectUnknown(call.arguments, ["sheetName"]);
      return fromHost(
        call.name,
        await host.getSheetVisibility(requireIdent(call.arguments, "sheetName")),
      );
    case "sheet.visibility.set":
      rejectUnknown(call.arguments, ["sheetName", "visibility"]);
      return fromHost(
        call.name,
        await host.setSheetVisibility(
          requireIdent(call.arguments, "sheetName"),
          requireVisibility(call.arguments),
        ),
      );
    case "sheet.protection.get":
      rejectUnknown(call.arguments, ["sheetName"]);
      return fromHost(
        call.name,
        await host.getSheetProtection(requireIdent(call.arguments, "sheetName")),
      );
    case "sheet.protection.protect":
      rejectUnknown(call.arguments, ["sheetName", "password"]);
      return fromHost(
        call.name,
        await host.protectSheet(
          requireIdent(call.arguments, "sheetName"),
          optionalPassword(call.arguments),
        ),
      );
    case "sheet.protection.unprotect":
      rejectUnknown(call.arguments, ["sheetName", "password"]);
      return fromHost(
        call.name,
        await host.unprotectSheet(
          requireIdent(call.arguments, "sheetName"),
          optionalPassword(call.arguments),
        ),
      );
    case "namedRange.list": {
      rejectUnknown(call.arguments, ["scope", "sheetName"]);
      if (call.arguments.scope == null || call.arguments.scope === "") {
        // Default workbook listing — sheetName without scope is invalid.
        if (
          Object.prototype.hasOwnProperty.call(call.arguments, "sheetName") &&
          call.arguments.sheetName != null &&
          call.arguments.sheetName !== ""
        ) {
          throw new Error("sheetName requires scope=worksheet (or omit sheetName for default workbook list)");
        }
        return fromHost(call.name, await host.listNamedRanges({}));
      }
      const scoped = resolveNamedScopeArgs(call.arguments);
      return fromHost(call.name, await host.listNamedRanges(scoped));
    }
    case "namedRange.create": {
      rejectUnknown(call.arguments, ["name", "refersTo", "scope", "sheetName", "visible"]);
      const scoped = resolveNamedScopeArgs(call.arguments);
      return fromHost(
        call.name,
        await host.createNamedRange({
          name: requireIdent(call.arguments, "name"),
          refersTo: requireValueString(call.arguments, "refersTo"),
          ...scoped,
          visible: optionalBoolean(call.arguments, "visible"),
        }),
      );
    }
    case "namedRange.update": {
      rejectUnknown(call.arguments, [
        "name",
        "scope",
        "sheetName",
        "newName",
        "refersTo",
        "visible",
      ]);
      const scoped = resolveNamedScopeArgs(call.arguments);
      return fromHost(
        call.name,
        await host.updateNamedRange({
          name: requireIdent(call.arguments, "name"),
          ...scoped,
          newName: optionalTrimmedName(call.arguments, "newName"),
          refersTo: optionalValueString(call.arguments, "refersTo"),
          visible: optionalBoolean(call.arguments, "visible"),
        }),
      );
    }
    case "namedRange.delete": {
      rejectUnknown(call.arguments, ["name", "scope", "sheetName"]);
      const scoped = resolveNamedScopeArgs(call.arguments);
      return fromHost(
        call.name,
        await host.deleteNamedRange({
          name: requireIdent(call.arguments, "name"),
          ...scoped,
        }),
      );
    }
    default:
      return null;
  }
}
