import type { HostAdapter } from "../host/types";
import { isChartAxisGroup } from "../host/chartAxisTypes";
import type { ChartAxisGroup } from "../host/chartAxisTypes";
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

function requirePositiveInt(args: Record<string, unknown>, key: string): number {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing argument: ${key}`);
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value) || !Number.isInteger(value) || value <= 0) {
    throw new Error(`${key} must be a positive integer`);
  }
  return value;
}

function requireAxisGroup(args: Record<string, unknown>): ChartAxisGroup {
  if (!Object.prototype.hasOwnProperty.call(args, "axisGroup")) {
    throw new Error("Missing argument: axisGroup");
  }
  if (args.axisGroup === undefined) throw new Error("axisGroup must not be undefined");
  if (args.axisGroup === null) throw new Error("axisGroup must not be null");
  if (!isChartAxisGroup(args.axisGroup)) {
    throw new Error("axisGroup must be primary|secondary");
  }
  return args.axisGroup;
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

export async function executeChartSeriesAxisGroupTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.axisGroup.update") return null;
  rejectUnknown(call.arguments, ["sheetName", "chartName", "seriesIndex", "axisGroup"]);
  return fromHost(
    call.name,
    await host.updateChartSeriesAxisGroup({
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
      axisGroup: requireAxisGroup(call.arguments),
    }),
  );
}
