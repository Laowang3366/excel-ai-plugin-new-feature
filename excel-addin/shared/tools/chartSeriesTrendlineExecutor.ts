import type { HostAdapter } from "../host/types";
import {
  isChartTrendlineType,
  type ChartTrendlineType,
} from "../host/chartSeriesTrendlineTypes";
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

function optionalPositiveInt(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value) || !Number.isInteger(value) || value <= 0) {
    throw new Error(`${key} must be a positive integer`);
  }
  return value;
}

function optionalNonNegFinite(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
    throw new Error(`${key} must be a non-negative finite number`);
  }
  return value;
}

/** number or "" (automatic intercept per Office.js). */
function optionalIntercept(args: Record<string, unknown>): number | "" | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "intercept")) return undefined;
  if (args.intercept === undefined) throw new Error("intercept must not be undefined");
  if (args.intercept === null) throw new Error("intercept must not be null");
  if (args.intercept === "") return "";
  if (typeof args.intercept !== "number" || !Number.isFinite(args.intercept)) {
    throw new Error("intercept must be a finite number or empty string for automatic");
  }
  return args.intercept;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalName(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "name")) return undefined;
  if (args.name === undefined) throw new Error("name must not be undefined");
  if (args.name === null) throw new Error("name must not be null");
  if (typeof args.name !== "string" || args.name.trim() === "") {
    throw new Error("name must be a non-empty string");
  }
  return args.name.trim();
}

function optionalType(args: Record<string, unknown>): ChartTrendlineType | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "type")) return undefined;
  if (args.type === undefined) throw new Error("type must not be undefined");
  if (args.type === null) throw new Error("type must not be null");
  if (!isChartTrendlineType(args.type)) {
    throw new Error("type must be a supported ChartTrendlineType");
  }
  return args.type;
}

function requireType(args: Record<string, unknown>): ChartTrendlineType {
  const t = optionalType(args);
  if (t === undefined) throw new Error("Missing argument: type");
  return t;
}

function rejectUnknown(args: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function parseFields(args: Record<string, unknown>, type?: ChartTrendlineType) {
  const fields = {
    type,
    name: optionalName(args),
    intercept: optionalIntercept(args),
    polynomialOrder: optionalPositiveInt(args, "polynomialOrder"),
    movingAveragePeriod: optionalPositiveInt(args, "movingAveragePeriod"),
    forwardPeriod: optionalNonNegFinite(args, "forwardPeriod"),
    backwardPeriod: optionalNonNegFinite(args, "backwardPeriod"),
    showEquation: optionalBoolean(args, "showEquation"),
    showRSquared: optionalBoolean(args, "showRSquared"),
  };
  if (fields.polynomialOrder !== undefined && fields.polynomialOrder < 2) {
    throw new Error("polynomialOrder must be >= 2");
  }
  if (fields.polynomialOrder !== undefined && fields.polynomialOrder > 6) {
    throw new Error("polynomialOrder must be <= 6");
  }
  if (fields.movingAveragePeriod !== undefined && fields.movingAveragePeriod < 2) {
    throw new Error("movingAveragePeriod must be >= 2");
  }
  const effectiveType = fields.type;
  if (fields.polynomialOrder !== undefined && effectiveType != null && effectiveType !== "polynomial") {
    throw new Error("polynomialOrder is only valid when type is polynomial");
  }
  if (
    fields.movingAveragePeriod !== undefined &&
    effectiveType != null &&
    effectiveType !== "movingAverage"
  ) {
    throw new Error("movingAveragePeriod is only valid when type is movingAverage");
  }
  return fields;
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

const FIELD_KEYS = [
  "type",
  "name",
  "intercept",
  "polynomialOrder",
  "movingAveragePeriod",
  "forwardPeriod",
  "backwardPeriod",
  "showEquation",
  "showRSquared",
];

export async function executeChartSeriesTrendlineTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (
    call.name !== "chart.series.trendlines.list" &&
    call.name !== "chart.series.trendlines.add" &&
    call.name !== "chart.series.trendlines.update" &&
    call.name !== "chart.series.trendlines.delete"
  ) {
    return null;
  }

  if (call.name === "chart.series.trendlines.list") {
    rejectUnknown(call.arguments, ["sheetName", "chartName", "seriesIndex"]);
    return fromHost(
      call.name,
      await host.listChartSeriesTrendlines(
        requireString(call.arguments, "sheetName"),
        requireString(call.arguments, "chartName"),
        requirePositiveInt(call.arguments, "seriesIndex"),
      ),
    );
  }

  if (call.name === "chart.series.trendlines.delete") {
    rejectUnknown(call.arguments, ["sheetName", "chartName", "seriesIndex", "trendlineIndex"]);
    return fromHost(
      call.name,
      await host.deleteChartSeriesTrendline(
        requireString(call.arguments, "sheetName"),
        requireString(call.arguments, "chartName"),
        requirePositiveInt(call.arguments, "seriesIndex"),
        requirePositiveInt(call.arguments, "trendlineIndex"),
      ),
    );
  }

  if (call.name === "chart.series.trendlines.add") {
    rejectUnknown(call.arguments, ["sheetName", "chartName", "seriesIndex", ...FIELD_KEYS]);
    const type = requireType(call.arguments);
    const fields = parseFields(call.arguments, type);
    return fromHost(
      call.name,
      await host.addChartSeriesTrendline({
        sheetName: requireString(call.arguments, "sheetName"),
        chartName: requireString(call.arguments, "chartName"),
        seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
        type,
        ...(fields.name !== undefined ? { name: fields.name } : {}),
        ...(fields.intercept !== undefined ? { intercept: fields.intercept } : {}),
        ...(fields.polynomialOrder !== undefined
          ? { polynomialOrder: fields.polynomialOrder }
          : {}),
        ...(fields.movingAveragePeriod !== undefined
          ? { movingAveragePeriod: fields.movingAveragePeriod }
          : {}),
        ...(fields.forwardPeriod !== undefined ? { forwardPeriod: fields.forwardPeriod } : {}),
        ...(fields.backwardPeriod !== undefined ? { backwardPeriod: fields.backwardPeriod } : {}),
        ...(fields.showEquation !== undefined ? { showEquation: fields.showEquation } : {}),
        ...(fields.showRSquared !== undefined ? { showRSquared: fields.showRSquared } : {}),
      }),
    );
  }

  // update
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "seriesIndex",
    "trendlineIndex",
    ...FIELD_KEYS,
  ]);
  const type = optionalType(call.arguments);
  const fields = parseFields(call.arguments, type);
  if (
    fields.type === undefined &&
    fields.name === undefined &&
    fields.intercept === undefined &&
    fields.polynomialOrder === undefined &&
    fields.movingAveragePeriod === undefined &&
    fields.forwardPeriod === undefined &&
    fields.backwardPeriod === undefined &&
    fields.showEquation === undefined &&
    fields.showRSquared === undefined
  ) {
    throw new Error("chart.series.trendlines.update requires at least one update field");
  }
  return fromHost(
    call.name,
    await host.updateChartSeriesTrendline({
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      seriesIndex: requirePositiveInt(call.arguments, "seriesIndex"),
      trendlineIndex: requirePositiveInt(call.arguments, "trendlineIndex"),
      ...(fields.type !== undefined ? { type: fields.type } : {}),
      ...(fields.name !== undefined ? { name: fields.name } : {}),
      ...(fields.intercept !== undefined ? { intercept: fields.intercept } : {}),
      ...(fields.polynomialOrder !== undefined
        ? { polynomialOrder: fields.polynomialOrder }
        : {}),
      ...(fields.movingAveragePeriod !== undefined
        ? { movingAveragePeriod: fields.movingAveragePeriod }
        : {}),
      ...(fields.forwardPeriod !== undefined ? { forwardPeriod: fields.forwardPeriod } : {}),
      ...(fields.backwardPeriod !== undefined ? { backwardPeriod: fields.backwardPeriod } : {}),
      ...(fields.showEquation !== undefined ? { showEquation: fields.showEquation } : {}),
      ...(fields.showRSquared !== undefined ? { showRSquared: fields.showRSquared } : {}),
    }),
  );
}
