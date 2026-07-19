import type { HostAdapter } from "../host/types";
import { isChartType } from "../host/types";
import type { ChartType, ToolCall, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";

function requireString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value;
}

function optionalString(args: Record<string, unknown>, key: string): string | undefined {
  const value = args[key];
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new Error(`Invalid string argument: ${key}`);
  return value;
}

function optionalFiniteNumber(args: Record<string, unknown>, key: string): number | undefined {
  if (!(key in args) || args[key] === undefined) return undefined;
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
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
  tool: Parameters<typeof mapHostResultToToolResult>[0],
  result: Parameters<typeof mapHostResultToToolResult>[1],
): ReturnType<typeof mapHostResultToToolResult> {
  return mapHostResultToToolResult(tool, result);
}

export async function executeChartTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "chart.list") {
    return fromHost(call.name, await host.listCharts(optionalString(call.arguments, "sheetName")));
  }

  if (call.name === "chart.create") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "sourceRange",
      "chartType",
      "name",
      "title",
      "left",
      "top",
      "width",
      "height",
    ]);
    return fromHost(
      call.name,
      await host.createChart({
        sheetName: requireString(call.arguments, "sheetName"),
        sourceRange: requireString(call.arguments, "sourceRange"),
        chartType: optionalChartType(call.arguments),
        name: optionalString(call.arguments, "name"),
        title: optionalString(call.arguments, "title"),
        left: optionalFiniteNumber(call.arguments, "left"),
        top: optionalFiniteNumber(call.arguments, "top"),
        width: optionalFiniteNumber(call.arguments, "width"),
        height: optionalFiniteNumber(call.arguments, "height"),
      }),
    );
  }

  if (call.name === "chart.delete") {
    return fromHost(
      call.name,
      await host.deleteChart(
        requireString(call.arguments, "sheetName"),
        requireString(call.arguments, "chartName"),
      ),
    );
  }

  return null;
}
