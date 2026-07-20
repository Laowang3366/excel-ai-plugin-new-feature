import {
  isChartLineStyle,
  type ChartTrendlineFormatUpdateInput,
} from "../host/chartSeriesTrendlineFormatTypes";
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

function optionalLineStyle(args: Record<string, unknown>): ChartTrendlineFormatUpdateInput["lineStyle"] {
  if (!Object.prototype.hasOwnProperty.call(args, "lineStyle")) return undefined;
  if (args.lineStyle === undefined) throw new Error("lineStyle must not be undefined");
  if (args.lineStyle === null) throw new Error("lineStyle must not be null");
  if (!isChartLineStyle(args.lineStyle)) {
    throw new Error(
      "lineStyle must be none|continuous|dash|dashDot|dashDotDot|dot|grey25|grey50|grey75|automatic|roundDot",
    );
  }
  return args.lineStyle;
}

function optionalWeight(args: Record<string, unknown>): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "weight")) return undefined;
  if (args.weight === undefined) throw new Error("weight must not be undefined");
  if (args.weight === null) throw new Error("weight must not be null");
  const value = args.weight;
  // Official ChartLineFormat.weight: finite number in points; no documented min/max.
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error("weight must be a finite number (points; host may reject invalid values)");
  }
  return value;
}

function optionalHexColor(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "color")) return undefined;
  if (args.color === undefined) throw new Error("color must not be undefined");
  if (args.color === null) throw new Error("color must not be null");
  if (typeof args.color !== "string") throw new Error("color must be a string (#RRGGBB)");
  const raw = args.color;
  if (raw === "") throw new Error("color must be #RRGGBB (empty string not allowed)");
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (!/^[0-9A-Fa-f]{6}$/.test(hex)) throw new Error("color must be #RRGGBB");
  return `#${hex.toUpperCase()}`;
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

export async function executeChartSeriesTrendlineFormatTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.trendlines.format.update") return null;
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "seriesIndex",
    "trendlineIndex",
    "color",
    "lineStyle",
    "weight",
  ]);
  const input: ChartTrendlineFormatUpdateInput = {
    sheetName: requireString(call.arguments, "sheetName"),
    chartName: requireString(call.arguments, "chartName"),
    seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
    trendlineIndex: requirePositiveInt(call.arguments, "trendlineIndex"),
    color: optionalHexColor(call.arguments),
    lineStyle: optionalLineStyle(call.arguments),
    weight: optionalWeight(call.arguments),
  };
  if (
    input.color === undefined &&
    input.lineStyle === undefined &&
    input.weight === undefined
  ) {
    throw new Error("chart.series.trendlines.format.update requires at least one update field");
  }
  return fromHost(call.name, await host.updateChartSeriesTrendlineFormat(input));
}
