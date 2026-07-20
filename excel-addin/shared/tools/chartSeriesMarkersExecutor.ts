import {
  isChartMarkerStyle,
  type ChartSeriesMarkersUpdateInput,
} from "../host/chartSeriesMarkersTypes";
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

function optionalMarkerStyle(args: Record<string, unknown>): ChartSeriesMarkersUpdateInput["markerStyle"] {
  if (!Object.prototype.hasOwnProperty.call(args, "markerStyle")) return undefined;
  if (args.markerStyle === undefined) throw new Error("markerStyle must not be undefined");
  if (args.markerStyle === null) throw new Error("markerStyle must not be null");
  if (!isChartMarkerStyle(args.markerStyle)) {
    throw new Error(
      "markerStyle must be automatic|none|square|diamond|triangle|x|star|dot|dash|circle|plus|picture",
    );
  }
  return args.markerStyle;
}

function optionalMarkerSize(args: Record<string, unknown>): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "markerSize")) return undefined;
  if (args.markerSize === undefined) throw new Error("markerSize must not be undefined");
  if (args.markerSize === null) throw new Error("markerSize must not be null");
  const value = args.markerSize;
  if (typeof value !== "number" || !Number.isInteger(value) || value < 2 || value > 72) {
    throw new Error("markerSize must be an integer from 2 to 72");
  }
  return value;
}

/** #RRGGBB only; missing omits; empty/null fail. */
function optionalHexColor(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "string") throw new Error(`${key} must be a string (#RRGGBB)`);
  const raw = args[key] as string;
  if (raw === "") throw new Error(`${key} must be #RRGGBB (empty string not allowed)`);
  const hex = raw.startsWith("#") ? raw.slice(1) : raw;
  if (!/^[0-9A-Fa-f]{6}$/.test(hex)) throw new Error(`${key} must be #RRGGBB`);
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

export async function executeChartSeriesMarkersTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.markers.update") return null;
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "seriesIndex",
    "markerStyle",
    "markerSize",
    "markerBackgroundColor",
    "markerForegroundColor",
  ]);
  const input: ChartSeriesMarkersUpdateInput = {
    sheetName: requireString(call.arguments, "sheetName"),
    chartName: requireString(call.arguments, "chartName"),
    seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
    markerStyle: optionalMarkerStyle(call.arguments),
    markerSize: optionalMarkerSize(call.arguments),
    markerBackgroundColor: optionalHexColor(call.arguments, "markerBackgroundColor"),
    markerForegroundColor: optionalHexColor(call.arguments, "markerForegroundColor"),
  };
  if (
    input.markerStyle === undefined &&
    input.markerSize === undefined &&
    input.markerBackgroundColor === undefined &&
    input.markerForegroundColor === undefined
  ) {
    throw new Error("chart.series.markers.update requires at least one update field");
  }
  return fromHost(call.name, await host.updateChartSeriesMarkers(input));
}
