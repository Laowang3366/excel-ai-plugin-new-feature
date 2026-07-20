import type {
  HostAdapter,
  RangeAutofitDirection,
  RangeDeleteShift,
  RangeInsertShift,
} from "../host/types";
import { mapHostResultToToolResult } from "./hostResultMapping";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing string argument: ${key}`);
  }
  const value = args[key];
  if (value === undefined) throw new Error(`${key} must not be undefined`);
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function requireEnum<T extends string>(
  args: Record<string, unknown>,
  key: string,
  values: readonly T[],
): T {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing argument: ${key}`);
  }
  const value = args[key];
  if (value === undefined) throw new Error(`${key} must not be undefined`);
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "string" || !values.includes(value as T)) {
    throw new Error(`${key} must be ${values.join("|")}`);
  }
  return value as T;
}

function rejectUnknown(args: Record<string, unknown>, allowed: readonly string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

export async function executeRangeStructureTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "range.insert") {
    rejectUnknown(call.arguments, ["sheetName", "range", "shift"]);
    return mapHostResultToToolResult(
      call.name,
      await host.insertRange({
        sheetName: requireString(call.arguments, "sheetName"),
        address: requireString(call.arguments, "range"),
        shift: requireEnum<RangeInsertShift>(call.arguments, "shift", ["down", "right"]),
      }),
    );
  }
  if (call.name === "range.delete") {
    rejectUnknown(call.arguments, ["sheetName", "range", "shift"]);
    return mapHostResultToToolResult(
      call.name,
      await host.deleteRange({
        sheetName: requireString(call.arguments, "sheetName"),
        address: requireString(call.arguments, "range"),
        shift: requireEnum<RangeDeleteShift>(call.arguments, "shift", ["up", "left"]),
      }),
    );
  }
  if (call.name === "range.autofit") {
    rejectUnknown(call.arguments, ["sheetName", "range", "direction"]);
    return mapHostResultToToolResult(
      call.name,
      await host.autofitRange({
        sheetName: requireString(call.arguments, "sheetName"),
        address: requireString(call.arguments, "range"),
        direction: requireEnum<RangeAutofitDirection>(call.arguments, "direction", [
          "rows",
          "columns",
          "both",
        ]),
      }),
    );
  }
  return null;
}
