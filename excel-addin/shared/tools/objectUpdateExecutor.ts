import type { ChartType, HostAdapter } from "../host/types";
import { isChartType } from "../host/types";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function optionalTrimmed(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] == null) return undefined;
  if (typeof args[key] !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = (args[key] as string).trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

function optionalResizeAddress(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "resizeAddress")) return undefined;
  if (args.resizeAddress === undefined) throw new Error("resizeAddress must not be undefined");
  if (args.resizeAddress === null) throw new Error("resizeAddress must not be null");
  return optionalTrimmed(args, "resizeAddress");
}

/** Chart title may be empty string to clear; do not trim content away. */
function optionalChartTitle(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "title") || args.title == null) return undefined;
  if (typeof args.title !== "string") throw new Error("Invalid string argument: title");
  return args.title;
}

/** Only missing / undefined omits; explicit null fails. */
function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) {
    return undefined;
  }
  if (typeof args[key] !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return args[key] as boolean;
}

function optionalFinite(args: Record<string, unknown>, key: string): number | undefined {
  const value = args[key];
  if (value == null) return undefined;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
  }
  return value;
}

function optionalPositive(args: Record<string, unknown>, key: string): number | undefined {
  const value = optionalFinite(args, key);
  if (value != null && value <= 0) throw new Error(`${key} must be > 0`);
  return value;
}

/** Chart.style: positive integer only; explicit null rejected. */
function optionalPositiveInt(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) return undefined;
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value) || !Number.isInteger(value) || value <= 0) {
    throw new Error(`${key} must be a positive integer`);
  }
  return value;
}

/** Key present with null/undefined/empty/wrong type fails; omit key only. */
function optionalChartType(args: Record<string, unknown>): ChartType | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "chartType")) return undefined;
  const value = args.chartType;
  if (value === undefined) throw new Error("chartType must not be undefined");
  if (value === null) throw new Error("chartType must not be null");
  if (typeof value !== "string" || value === "" || !isChartType(value)) {
    throw new Error(
      "chartType must be column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers",
    );
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

export async function executeObjectUpdateTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "table.update") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "tableName",
      "newName",
      "style",
      "showHeaders",
      "showTotals",
      "showFilterButton",
      "showBandedRows",
      "showBandedColumns",
      "showFirstColumn",
      "showLastColumn",
      "resizeAddress",
    ]);
    const input = {
      sheetName: requireString(call.arguments, "sheetName"),
      tableName: requireString(call.arguments, "tableName"),
      newName: optionalTrimmed(call.arguments, "newName"),
      style: optionalTrimmed(call.arguments, "style"),
      showHeaders: optionalBoolean(call.arguments, "showHeaders"),
      showTotals: optionalBoolean(call.arguments, "showTotals"),
      showFilterButton: optionalBoolean(call.arguments, "showFilterButton"),
      showBandedRows: optionalBoolean(call.arguments, "showBandedRows"),
      showBandedColumns: optionalBoolean(call.arguments, "showBandedColumns"),
      showFirstColumn: optionalBoolean(call.arguments, "showFirstColumn"),
      showLastColumn: optionalBoolean(call.arguments, "showLastColumn"),
      resizeAddress: optionalResizeAddress(call.arguments),
    };
    if (
      input.newName == null &&
      input.style == null &&
      input.showHeaders == null &&
      input.showTotals == null &&
      input.showFilterButton == null &&
      input.showBandedRows == null &&
      input.showBandedColumns == null &&
      input.showFirstColumn == null &&
      input.showLastColumn == null &&
      input.resizeAddress == null
    ) {
      throw new Error("table.update requires at least one update field");
    }
    return fromHost(call.name, await host.updateTable(input));
  }

  if (call.name === "chart.update") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "chartName",
      "newName",
      "chartType",
      "title",
      "showTitle",
      "style",
      "showLegend",
      "left",
      "top",
      "width",
      "height",
    ]);
    const input = {
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      newName: optionalTrimmed(call.arguments, "newName"),
      chartType: optionalChartType(call.arguments),
      title: optionalChartTitle(call.arguments),
      showTitle: optionalBoolean(call.arguments, "showTitle"),
      style: optionalPositiveInt(call.arguments, "style"),
      showLegend: optionalBoolean(call.arguments, "showLegend"),
      left: optionalFinite(call.arguments, "left"),
      top: optionalFinite(call.arguments, "top"),
      width: optionalPositive(call.arguments, "width"),
      height: optionalPositive(call.arguments, "height"),
    };
    if (
      input.newName == null &&
      input.chartType == null &&
      input.title == null &&
      input.showTitle == null &&
      input.style == null &&
      input.showLegend == null &&
      input.left == null &&
      input.top == null &&
      input.width == null &&
      input.height == null
    ) {
      throw new Error("chart.update requires at least one update field");
    }
    return fromHost(call.name, await host.updateChart(input));
  }

  return null;
}
