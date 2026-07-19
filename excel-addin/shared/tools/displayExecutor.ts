import type { HostAdapter, SheetDisplayUpdateInput } from "../host/types";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

/** Only missing / undefined omits; explicit null fails. */
function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

/** Empty string = automatic; otherwise normalize to #RRGGBB uppercase. */
export function normalizeTabColor(raw: unknown): string {
  if (typeof raw !== "string") throw new Error("tabColor must be a string");
  if (raw === "") return "";
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (!/^[0-9A-Fa-f]{6}$/.test(hex)) {
    throw new Error("tabColor must be empty or six-digit hex (#RRGGBB)");
  }
  return `#${hex.toUpperCase()}`;
}

/** Only missing / undefined omits; explicit null fails. */
function optionalTabColor(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "tabColor") || args.tabColor === undefined) {
    return undefined;
  }
  return normalizeTabColor(args.tabColor);
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

export async function executeDisplayTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "sheet.display.get") {
    rejectUnknown(call.arguments, ["sheetName"]);
    return fromHost(
      call.name,
      await host.getSheetDisplay(requireString(call.arguments, "sheetName")),
    );
  }
  if (call.name === "sheet.display.set") {
    rejectUnknown(call.arguments, ["sheetName", "tabColor", "showGridlines", "showHeadings"]);
    const input: SheetDisplayUpdateInput = {
      sheetName: requireString(call.arguments, "sheetName"),
      tabColor: optionalTabColor(call.arguments),
      showGridlines: optionalBoolean(call.arguments, "showGridlines"),
      showHeadings: optionalBoolean(call.arguments, "showHeadings"),
    };
    if (
      input.tabColor === undefined &&
      input.showGridlines === undefined &&
      input.showHeadings === undefined
    ) {
      throw new Error("sheet.display.set requires at least one update field");
    }
    return fromHost(call.name, await host.setSheetDisplay(input));
  }
  return null;
}
