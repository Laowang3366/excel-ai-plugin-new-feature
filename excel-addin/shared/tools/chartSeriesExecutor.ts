import type { ChartType, HostAdapter } from "../host/types";
import { isChartType } from "../host/types";
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

function optionalTrimmed(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    return undefined;
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = (args[key] as string).trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    return undefined;
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalChartType(args: Record<string, unknown>): ChartType | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "chartType")) {
    return undefined;
  }
  if (args.chartType === undefined) throw new Error("chartType must not be undefined");
  if (args.chartType === null) throw new Error("chartType must not be null");
  if (!isChartType(args.chartType)) {
    throw new Error(
      "chartType must be column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers",
    );
  }
  return args.chartType;
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

export async function executeChartSeriesTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "chart.series.list") {
    rejectUnknown(call.arguments, ["sheetName", "chartName"]);
    return fromHost(
      call.name,
      await host.listChartSeries(
        requireString(call.arguments, "sheetName"),
        requireString(call.arguments, "chartName"),
      ),
    );
  }

  if (call.name === "chart.series.update") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "chartName",
      "seriesIndex",
      "newName",
      "chartType",
      "smooth",
    ]);
    const input = {
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
      newName: optionalTrimmed(call.arguments, "newName"),
      chartType: optionalChartType(call.arguments),
      smooth: optionalBoolean(call.arguments, "smooth"),
    };
    if (input.newName == null && input.chartType == null && input.smooth == null) {
      throw new Error("chart.series.update requires at least one update field");
    }
    return fromHost(call.name, await host.updateChartSeries(input));
  }

  return null;
}
