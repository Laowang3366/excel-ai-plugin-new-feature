import type { HostAdapter } from "../host/types";
import { isChartType } from "../host/types";
import type { ChartType, ToolCall, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";
import {
  optionalIdent,
  rejectUnknownFields,
  requireIdent,
} from "./argValidation";

/** Title is clearable/special: keep raw spaces; null/"" omit. */
function optionalChartTitle(args: Record<string, unknown>): string | undefined {
  const value = args.title;
  if (value == null || value === "") return undefined;
  if (typeof value !== "string") throw new Error("Invalid string argument: title");
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
  if (typeof value !== "string" || value === "") {
    throw new Error(
      "chartType must be column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers",
    );
  }
  const trimmed = value.trim();
  if (!isChartType(trimmed)) {
    throw new Error(
      "chartType must be column|line|bar|area|pie|scatter|doughnut|bubble|radar|linemarkers",
    );
  }
  return trimmed;
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
    rejectUnknownFields(call.arguments, ["sheetName"]);
    return fromHost(call.name, await host.listCharts(optionalIdent(call.arguments, "sheetName")));
  }

  if (call.name === "chart.create") {
    rejectUnknownFields(call.arguments, [
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
        sheetName: requireIdent(call.arguments, "sheetName"),
        sourceRange: requireIdent(call.arguments, "sourceRange"),
        chartType: optionalChartType(call.arguments),
        name: optionalIdent(call.arguments, "name"),
        title: optionalChartTitle(call.arguments),
        left: optionalFiniteNumber(call.arguments, "left"),
        top: optionalFiniteNumber(call.arguments, "top"),
        width: optionalFiniteNumber(call.arguments, "width"),
        height: optionalFiniteNumber(call.arguments, "height"),
      }),
    );
  }

  if (call.name === "chart.delete") {
    rejectUnknownFields(call.arguments, ["sheetName", "chartName"]);
    return fromHost(
      call.name,
      await host.deleteChart(
        requireIdent(call.arguments, "sheetName"),
        requireIdent(call.arguments, "chartName"),
      ),
    );
  }

  return null;
}
