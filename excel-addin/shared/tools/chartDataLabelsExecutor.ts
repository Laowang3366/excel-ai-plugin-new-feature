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

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalNumberFormat(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "numberFormat")) return undefined;
  if (args.numberFormat === undefined) throw new Error("numberFormat must not be undefined");
  if (args.numberFormat === null) throw new Error("numberFormat must not be null");
  if (typeof args.numberFormat !== "string") {
    throw new Error("Invalid string argument: numberFormat");
  }
  const trimmed = args.numberFormat.trim();
  if (trimmed === "") throw new Error("numberFormat must be non-empty");
  return trimmed;
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

export async function executeChartDataLabelsTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.series.dataLabels.update") return null;
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "seriesIndex",
    "enabled",
    "showValue",
    "showCategoryName",
    "showSeriesName",
    "numberFormat",
  ]);
  const input = {
    sheetName: requireString(call.arguments, "sheetName"),
    chartName: requireString(call.arguments, "chartName"),
    seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
    enabled: optionalBoolean(call.arguments, "enabled"),
    showValue: optionalBoolean(call.arguments, "showValue"),
    showCategoryName: optionalBoolean(call.arguments, "showCategoryName"),
    showSeriesName: optionalBoolean(call.arguments, "showSeriesName"),
    numberFormat: optionalNumberFormat(call.arguments),
  };
  if (
    input.enabled === undefined &&
    input.showValue === undefined &&
    input.showCategoryName === undefined &&
    input.showSeriesName === undefined &&
    input.numberFormat === undefined
  ) {
    throw new Error("chart.series.dataLabels.update requires at least one update field");
  }
  const hasOtherLabelFields =
    input.showValue !== undefined ||
    input.showCategoryName !== undefined ||
    input.showSeriesName !== undefined ||
    input.numberFormat !== undefined;
  if (input.enabled === false && hasOtherLabelFields) {
    throw new Error(
      "enabled=false cannot be combined with showValue/showCategoryName/showSeriesName/numberFormat",
    );
  }
  return fromHost(call.name, await host.updateChartDataLabels(input));
}
