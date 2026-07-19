import type { HostAdapter } from "../host/types";
import {
  isChartAxisGroup,
  isChartAxisKind,
  type ChartAxisGroup,
  type ChartAxisKind,
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

function optionalMajorUnit(args: Record<string, unknown>): number | undefined {
  const value = optionalFinite(args, "majorUnit");
  if (value != null && value < 0) throw new Error("majorUnit must be >= 0");
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
    "title",
    "minimum",
    "maximum",
    "majorUnit",
    "numberFormat",
    "reverse",
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
    numberFormat: optionalNumberFormat(call.arguments),
    reverse: optionalBoolean(call.arguments, "reverse"),
  };
  if (
    input.title === undefined &&
    input.minimum === undefined &&
    input.maximum === undefined &&
    input.majorUnit === undefined &&
    input.numberFormat === undefined &&
    input.reverse === undefined
  ) {
    throw new Error("chart.axes.update requires at least one update field");
  }
  return fromHost(call.name, await host.updateChartAxis(input));
}
