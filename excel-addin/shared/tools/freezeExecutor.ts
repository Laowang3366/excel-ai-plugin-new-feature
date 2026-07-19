import type { HostAdapter, SheetFreezeCommand, SheetFreezeSetInput } from "../host/types";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (typeof args[key] !== "string" || (args[key] as string).trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return (args[key] as string).trim();
}

function requireCommand(args: Record<string, unknown>): SheetFreezeCommand {
  if (!Object.prototype.hasOwnProperty.call(args, "command") || args.command === undefined) {
    throw new Error("Missing command");
  }
  const value = args.command;
  if (value !== "rows" && value !== "columns" && value !== "at" && value !== "clear") {
    throw new Error("command must be rows|columns|at|clear");
  }
  return value;
}

function requirePositiveCount(args: Record<string, unknown>): number {
  if (!Object.prototype.hasOwnProperty.call(args, "count") || args.count === undefined) {
    throw new Error("count is required");
  }
  if (typeof args.count !== "number" || !Number.isInteger(args.count) || args.count < 1) {
    throw new Error("count must be a positive integer");
  }
  return args.count;
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

export async function executeFreezeTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "sheet.freeze.get") {
    rejectUnknown(call.arguments, ["sheetName"]);
    return fromHost(
      call.name,
      await host.getSheetFreeze(requireString(call.arguments, "sheetName")),
    );
  }
  if (call.name === "sheet.freeze.set") {
    const command = requireCommand(call.arguments);
    if (command === "rows" || command === "columns") {
      rejectUnknown(call.arguments, ["sheetName", "command", "count"]);
      const input: SheetFreezeSetInput = {
        sheetName: requireString(call.arguments, "sheetName"),
        command,
        count: requirePositiveCount(call.arguments),
      };
      return fromHost(call.name, await host.setSheetFreeze(input));
    }
    if (command === "at") {
      rejectUnknown(call.arguments, ["sheetName", "command", "address"]);
      const input: SheetFreezeSetInput = {
        sheetName: requireString(call.arguments, "sheetName"),
        command,
        address: requireString(call.arguments, "address"),
      };
      return fromHost(call.name, await host.setSheetFreeze(input));
    }
    rejectUnknown(call.arguments, ["sheetName", "command"]);
    return fromHost(
      call.name,
      await host.setSheetFreeze({
        sheetName: requireString(call.arguments, "sheetName"),
        command: "clear",
      }),
    );
  }
  return null;
}
