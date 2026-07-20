import type { HostAdapter } from "../host/types";
import {
  isChartAxisDisplayUnit,
  isChartAxisGroup,
  isChartAxisKind,
  isChartAxisPosition,
  isChartAxisScaleType,
  isChartAxisTickLabelPosition,
  isChartAxisTickMark,
  type ChartAxisDisplayUnit,
  type ChartAxisGroup,
  type ChartAxisKind,
  type ChartAxisPosition,
  type ChartAxisScaleType,
  type ChartAxisTickLabelPosition,
  type ChartAxisTickMark,
} from "../host/chartAxisTypes";
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

function requireKind(args: Record<string, unknown>): ChartAxisKind {
  if (!Object.prototype.hasOwnProperty.call(args, "kind")) {
    throw new Error("Missing argument: kind");
  }
  if (args.kind === undefined) throw new Error("kind must not be undefined");
  if (args.kind === null) throw new Error("kind must not be null");
  if (!isChartAxisKind(args.kind)) throw new Error("kind must be category|value");
  return args.kind;
}

function optionalGroup(args: Record<string, unknown>): ChartAxisGroup | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "group")) return undefined;
  if (args.group === undefined) throw new Error("group must not be undefined");
  if (args.group === null) throw new Error("group must not be null");
  if (!isChartAxisGroup(args.group)) throw new Error("group must be primary|secondary");
  return args.group;
}

/** title: only "" clears; whitespace-only rejected; otherwise keep content (may include edge spaces). */
function optionalTitle(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "title")) return undefined;
  if (args.title === undefined) throw new Error("title must not be undefined");
  if (args.title === null) throw new Error("title must not be null");
  if (typeof args.title !== "string") throw new Error("Invalid string argument: title");
  if (args.title === "") return "";
  if (args.title.trim() === "") throw new Error("title must be empty or non-whitespace");
  return args.title;
}

function optionalFinite(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
  }
  return value;
}

function optionalPositiveFinite(args: Record<string, unknown>, key: string): number | undefined {
  const value = optionalFinite(args, key);
  if (value != null && value <= 0) throw new Error(`${key} must be a positive finite number`);
  return value;
}

/** majorUnit: finite number (>=0 keeps prior contract) or "" for automatic. */
function optionalMajorUnit(args: Record<string, unknown>): number | "" | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "majorUnit")) return undefined;
  if (args.majorUnit === undefined) throw new Error("majorUnit must not be undefined");
  if (args.majorUnit === null) throw new Error("majorUnit must not be null");
  if (args.majorUnit === "") return "";
  const value = args.majorUnit;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error("majorUnit must be a finite number or empty string");
  }
  if (value < 0) throw new Error("majorUnit must be >= 0");
  return value;
}

/** minorUnit: finite number or "" for automatic; no invented range bounds. */
function optionalMinorUnit(args: Record<string, unknown>): number | "" | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "minorUnit")) return undefined;
  if (args.minorUnit === undefined) throw new Error("minorUnit must not be undefined");
  if (args.minorUnit === null) throw new Error("minorUnit must not be null");
  if (args.minorUnit === "") return "";
  const value = args.minorUnit;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error("minorUnit must be a finite number or empty string");
  }
  return value;
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

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalDisplayUnit(args: Record<string, unknown>): ChartAxisDisplayUnit | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "displayUnit")) return undefined;
  if (args.displayUnit === undefined) throw new Error("displayUnit must not be undefined");
  if (args.displayUnit === null) throw new Error("displayUnit must not be null");
  if (!isChartAxisDisplayUnit(args.displayUnit)) {
    throw new Error("displayUnit must be a supported ChartAxisDisplayUnit");
  }
  return args.displayUnit;
}

function optionalScaleType(args: Record<string, unknown>): ChartAxisScaleType | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "scaleType")) return undefined;
  if (args.scaleType === undefined) throw new Error("scaleType must not be undefined");
  if (args.scaleType === null) throw new Error("scaleType must not be null");
  if (!isChartAxisScaleType(args.scaleType)) {
    throw new Error("scaleType must be linear|logarithmic");
  }
  return args.scaleType;
}

function optionalTickMark(
  args: Record<string, unknown>,
  key: "majorTickMark" | "minorTickMark",
): ChartAxisTickMark | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (!isChartAxisTickMark(args[key])) {
    throw new Error(`${key} must be none|cross|inside|outside`);
  }
  return args[key] as ChartAxisTickMark;
}

function optionalTickLabelPosition(
  args: Record<string, unknown>,
): ChartAxisTickLabelPosition | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "tickLabelPosition")) return undefined;
  if (args.tickLabelPosition === undefined) {
    throw new Error("tickLabelPosition must not be undefined");
  }
  if (args.tickLabelPosition === null) throw new Error("tickLabelPosition must not be null");
  if (!isChartAxisTickLabelPosition(args.tickLabelPosition)) {
    throw new Error("tickLabelPosition must be nextToAxis|high|low|none");
  }
  return args.tickLabelPosition;
}

function optionalPosition(args: Record<string, unknown>): ChartAxisPosition | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "position")) return undefined;
  if (args.position === undefined) throw new Error("position must not be undefined");
  if (args.position === null) throw new Error("position must not be null");
  if (!isChartAxisPosition(args.position)) {
    throw new Error("position must be automatic|maximum|minimum|custom");
  }
  return args.position;
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

const UPDATE_FIELDS = [
  "title",
  "minimum",
  "maximum",
  "majorUnit",
  "minorUnit",
  "numberFormat",
  "reverse",
  "displayUnit",
  "customDisplayUnit",
  "scaleType",
  "logBase",
  "showDisplayUnitLabel",
  "majorGridlinesVisible",
  "minorGridlinesVisible",
  "majorTickMark",
  "minorTickMark",
  "tickLabelPosition",
  "position",
  "positionAt",
  "linkNumberFormat",
] as const;

export async function executeChartAxesTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.axes.update") return null;
  rejectUnknown(call.arguments, [
    "sheetName",
    "chartName",
    "kind",
    "group",
    ...UPDATE_FIELDS,
  ]);
  const input = {
    sheetName: requireString(call.arguments, "sheetName"),
    chartName: requireString(call.arguments, "chartName"),
    kind: requireKind(call.arguments),
    group: optionalGroup(call.arguments),
    title: optionalTitle(call.arguments),
    minimum: optionalFinite(call.arguments, "minimum"),
    maximum: optionalFinite(call.arguments, "maximum"),
    majorUnit: optionalMajorUnit(call.arguments),
    minorUnit: optionalMinorUnit(call.arguments),
    numberFormat: optionalNumberFormat(call.arguments),
    reverse: optionalBoolean(call.arguments, "reverse"),
    displayUnit: optionalDisplayUnit(call.arguments),
    customDisplayUnit: optionalPositiveFinite(call.arguments, "customDisplayUnit"),
    scaleType: optionalScaleType(call.arguments),
    logBase: optionalPositiveFinite(call.arguments, "logBase"),
    showDisplayUnitLabel: optionalBoolean(call.arguments, "showDisplayUnitLabel"),
    majorGridlinesVisible: optionalBoolean(call.arguments, "majorGridlinesVisible"),
    minorGridlinesVisible: optionalBoolean(call.arguments, "minorGridlinesVisible"),
    majorTickMark: optionalTickMark(call.arguments, "majorTickMark"),
    minorTickMark: optionalTickMark(call.arguments, "minorTickMark"),
    tickLabelPosition: optionalTickLabelPosition(call.arguments),
    position: optionalPosition(call.arguments),
    positionAt: optionalFinite(call.arguments, "positionAt"),
    linkNumberFormat: optionalBoolean(call.arguments, "linkNumberFormat"),
  };
  if (input.displayUnit === "custom" && input.customDisplayUnit === undefined) {
    throw new Error("customDisplayUnit is required when displayUnit is custom");
  }
  if (
    input.customDisplayUnit !== undefined &&
    input.displayUnit !== undefined &&
    input.displayUnit !== "custom"
  ) {
    throw new Error("customDisplayUnit requires displayUnit custom (or omit displayUnit)");
  }
  if (input.position === "custom" && input.positionAt === undefined) {
    throw new Error("positionAt is required when position is custom");
  }
  if (!UPDATE_FIELDS.some((key) => input[key] !== undefined)) {
    throw new Error("chart.axes.update requires at least one update field");
  }
  return fromHost(call.name, await host.updateChartAxis(input));
}
