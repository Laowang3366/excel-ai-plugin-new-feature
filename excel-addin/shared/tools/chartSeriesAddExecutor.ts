import type { HostAdapter } from "../host/types";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function optionalName(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "name")) return undefined;
  if (args.name === undefined) throw new Error("name must not be undefined");
  if (args.name === null) throw new Error("name must not be null");
  if (typeof args.name !== "string" || args.name.trim() === "") {
    throw new Error("name must be a non-empty string when provided");
  }
  return args.name.trim();
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

export async function executeChartSeriesAddTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.add") return null;
  rejectUnknown(call.arguments, ["sheetName", "chartName", "name"]);
  const name = optionalName(call.arguments);
  return fromHost(
    call.name,
    await host.addChartSeries({
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      ...(name != null ? { name } : {}),
    }),
  );
}
