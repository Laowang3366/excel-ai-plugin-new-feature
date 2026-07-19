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

function optionalRange(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "string" || (args[key] as string).trim() === "") {
    throw new Error(`${key} must be a non-empty string when provided`);
  }
  return (args[key] as string).trim();
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

export async function executeChartSeriesValuesTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.values.update") return null;
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "seriesIndex",
    "valuesRange",
    "xValuesRange",
  ]);
  const valuesRange = optionalRange(call.arguments, "valuesRange");
  const xValuesRange = optionalRange(call.arguments, "xValuesRange");
  if (valuesRange == null && xValuesRange == null) {
    throw new Error("chart.series.values.update requires valuesRange and/or xValuesRange");
  }
  return fromHost(
    call.name,
    await host.updateChartSeriesValues({
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
      ...(valuesRange != null ? { valuesRange } : {}),
      ...(xValuesRange != null ? { xValuesRange } : {}),
    }),
  );
}
