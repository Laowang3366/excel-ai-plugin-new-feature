import type { HostAdapter, RangeExpandMode, RangeFormat } from "../host/types";
import type { CellValue, ToolCall, ToolFailure, ToolName, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";
import {
  rejectUnknownCoreToolArguments,
  rejectUnknownRangeFormatFields,
} from "./argValidation";
import { requireCfRule, requireDvRule } from "./ruleValidation";
import { executeSheetOperation } from "./sheetOperation";
import { executeDisplayTool } from "./displayExecutor";
import { executeFreezeTool } from "./freezeExecutor";
import { executeChartTool } from "./chartExecutor";
import { executeChartSeriesTool } from "./chartSeriesExecutor";
import { executeChartAxesTool } from "./chartAxesExecutor";
import { executeChartDataLabelsTool } from "./chartDataLabelsExecutor";
import { executeChartSeriesAxisGroupTool } from "./chartSeriesAxisGroupExecutor";
import { executeChartSeriesAddTool } from "./chartSeriesAddExecutor";
import { executeChartSeriesDeleteTool } from "./chartSeriesDeleteExecutor";
import { executeChartSeriesValuesTool } from "./chartSeriesValuesExecutor";
import { executeChartSeriesBubbleSizesTool } from "./chartSeriesBubbleSizesExecutor";
import { executeChartImageTool } from "./chartImageExecutor";
import { executeRangeImageTool } from "./rangeImageExecutor";
import { executeChartSourceTool } from "./chartSourceExecutor";
import { executeObjectUpdateTool } from "./objectUpdateExecutor";
import { executePageLayoutTool } from "./pageLayoutExecutor";
import { executeShapeTool } from "./shapeExecutor";
import { executeStructureTool } from "./structureExecutor";
import { executeTableUnlistTool } from "./tableUnlistExecutor";
import { writeFormulaWithVerify, writeRangeWithVerify } from "./writeWithVerify";

function fail(tool: ToolName, error: string, detail?: unknown): ToolFailure {
  return { ok: false, tool, error, detail };
}

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

function requireValues(args: Record<string, unknown>): CellValue[][] {
  const values = args.values;
  if (!Array.isArray(values) || values.length === 0) {
    throw new Error("values must be a non-empty 2D array");
  }
  return values as CellValue[][];
}

function requireFormat(args: Record<string, unknown>): RangeFormat {
  const format = args.format;
  if (!format || typeof format !== "object" || Array.isArray(format)) {
    throw new Error("format must be an object");
  }
  rejectUnknownRangeFormatFields(format as Record<string, unknown>);
  return format as RangeFormat;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!(key in args) || args[key] === undefined) return undefined;
  if (typeof args[key] !== "boolean") {
    throw new Error(`${key} must be a boolean`);
  }
  return args[key] as boolean;
}

function optionalFiniteNumber(args: Record<string, unknown>, key: string): number | undefined {
  if (!(key in args) || args[key] === undefined) return undefined;
  const value = args[key];
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
  }
  return value;
}

function optionalExpand(args: Record<string, unknown>): RangeExpandMode | undefined {
  if (!("expand" in args) || args.expand === undefined || args.expand === "") {
    return undefined;
  }
  const value = args.expand;
  if (
    value !== "none" &&
    value !== "spill" &&
    value !== "currentArray" &&
    value !== "currentRegion"
  ) {
    throw new Error("expand must be none|spill|currentArray|currentRegion");
  }
  return value;
}

function requireSheetOperation(args: Record<string, unknown>): "add" | "rename" | "delete" | "copy" | "move" {
  const value = args.operation;
  if (
    value !== "add" &&
    value !== "rename" &&
    value !== "delete" &&
    value !== "copy" &&
    value !== "move"
  ) {
    throw new Error("operation must be add|rename|delete|copy|move");
  }
  return value;
}

function fromHost(
  tool: Parameters<typeof mapHostResultToToolResult>[0],
  result: Parameters<typeof mapHostResultToToolResult>[1],
): ReturnType<typeof mapHostResultToToolResult> {
  return mapHostResultToToolResult(tool, result);
}

export class ToolExecutor {
  constructor(private readonly host: HostAdapter) {}

  async execute(call: ToolCall): Promise<ToolResult> {
    try {
      rejectUnknownCoreToolArguments(call.name, call.arguments);
      switch (call.name) {
        case "host.status":
          return fromHost(call.name, await this.host.getStatus());
        case "selection.get":
          return fromHost(call.name, await this.host.getSelection());
        case "range.read":
          return fromHost(
            call.name,
            await this.host.readRange(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
              optionalExpand(call.arguments),
            ),
          );
        case "range.write":
          return await writeRangeWithVerify(this.host, {
            sheetName: requireString(call.arguments, "sheetName"),
            range: requireString(call.arguments, "range"),
            values: requireValues(call.arguments),
            verify: call.arguments.verify !== false,
          });
        case "range.clear":
          return fromHost(
            call.name,
            await this.host.clearRange(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
            ),
          );
        case "range.format.read":
          return fromHost(
            call.name,
            await this.host.readFormat(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
            ),
          );
        case "range.format.write":
          return fromHost(
            call.name,
            await this.host.writeFormat(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
              requireFormat(call.arguments),
            ),
          );
        case "formula.read": {
          const read = await this.host.readRange(
            requireString(call.arguments, "sheetName"),
            requireString(call.arguments, "range"),
          );
          if (!read.ok) return fromHost(call.name, read);
          return {
            ok: true,
            tool: call.name,
            data: {
              sheetName: read.data.sheetName,
              address: read.data.address,
              formulas: read.data.formulas,
            },
          };
        }
        case "formula.write":
          return await writeFormulaWithVerify(this.host, {
            sheetName: requireString(call.arguments, "sheetName"),
            range: requireString(call.arguments, "range"),
            formula: requireString(call.arguments, "formula"),
            verify: call.arguments.verify !== false,
          });
        case "formula.context":
          return fromHost(
            call.name,
            await this.host.getFormulaContext(
              requireString(call.arguments, "sheetName"),
              optionalString(call.arguments, "range"),
            ),
          );
        case "sheet.list":
          return fromHost(call.name, await this.host.listSheets());
        case "sheet.add":
          return fromHost(
            call.name,
            await this.host.addSheet(requireString(call.arguments, "sheetName")),
          );
        case "sheet.rename":
          return fromHost(
            call.name,
            await this.host.renameSheet(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "newName"),
            ),
          );
        case "sheet.delete":
          return fromHost(
            call.name,
            await this.host.deleteSheet(requireString(call.arguments, "sheetName")),
          );
        case "sheet.operation":
          return await executeSheetOperation(this.host, call.arguments, fromHost, {
            requireString,
            optionalString,
            optionalFiniteNumber,
            requireSheetOperation,
          });
        case "table.list":
          return fromHost(
            call.name,
            await this.host.listTables(optionalString(call.arguments, "sheetName")),
          );
        case "table.create":
          return fromHost(
            call.name,
            await this.host.createTable({
              sheetName: requireString(call.arguments, "sheetName"),
              address: requireString(call.arguments, "range"),
              name: optionalString(call.arguments, "name"),
              hasHeaders: optionalBoolean(call.arguments, "hasHeaders"),
            }),
          );
        case "table.delete":
          return fromHost(
            call.name,
            await this.host.deleteTable(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "tableName"),
            ),
          );
        case "workbook.inspect":
          return fromHost(call.name, await this.host.inspectWorkbook());
        case "conditionalFormat.list":
          return fromHost(
            call.name,
            await this.host.listConditionalFormats(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
            ),
          );
        case "conditionalFormat.add":
          return fromHost(
            call.name,
            await this.host.addConditionalFormat({
              sheetName: requireString(call.arguments, "sheetName"),
              range: requireString(call.arguments, "range"),
              rule: requireCfRule(call.arguments),
            }),
          );
        case "conditionalFormat.delete":
          return fromHost(
            call.name,
            await this.host.deleteConditionalFormat(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
              requireString(call.arguments, "id"),
            ),
          );
        case "dataValidation.read":
          return fromHost(
            call.name,
            await this.host.readDataValidation(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
            ),
          );
        case "dataValidation.write":
          return fromHost(
            call.name,
            await this.host.writeDataValidation({
              sheetName: requireString(call.arguments, "sheetName"),
              range: requireString(call.arguments, "range"),
              rule: requireDvRule(call.arguments),
            }),
          );
        case "dataValidation.clear":
          return fromHost(
            call.name,
            await this.host.clearDataValidation(
              requireString(call.arguments, "sheetName"),
              requireString(call.arguments, "range"),
            ),
          );
        default: {
          const chart = await executeChartTool(this.host, call);
          if (chart) return chart;
          const structure = await executeStructureTool(this.host, call);
          if (structure) return structure;
          const objectUpdate = await executeObjectUpdateTool(this.host, call);
          if (objectUpdate) return objectUpdate;
          const chartSeries = await executeChartSeriesTool(this.host, call);
          if (chartSeries) return chartSeries;
          const chartSource = await executeChartSourceTool(this.host, call);
          if (chartSource) return chartSource;
          const chartAxes = await executeChartAxesTool(this.host, call);
          if (chartAxes) return chartAxes;
          const chartDataLabels = await executeChartDataLabelsTool(this.host, call);
          if (chartDataLabels) return chartDataLabels;
          const chartSeriesAxisGroup = await executeChartSeriesAxisGroupTool(this.host, call);
          if (chartSeriesAxisGroup) return chartSeriesAxisGroup;
          const chartSeriesDelete = await executeChartSeriesDeleteTool(this.host, call);
          if (chartSeriesDelete) return chartSeriesDelete;
          const chartSeriesAdd = await executeChartSeriesAddTool(this.host, call);
          if (chartSeriesAdd) return chartSeriesAdd;
          const chartSeriesValues = await executeChartSeriesValuesTool(this.host, call);
          if (chartSeriesValues) return chartSeriesValues;
          const chartSeriesBubbleSizes = await executeChartSeriesBubbleSizesTool(this.host, call);
          if (chartSeriesBubbleSizes) return chartSeriesBubbleSizes;
          const chartImage = await executeChartImageTool(this.host, call);
          if (chartImage) return chartImage;
          const rangeImage = await executeRangeImageTool(this.host, call);
          if (rangeImage) return rangeImage;
          const display = await executeDisplayTool(this.host, call);
          if (display) return display;
          const freeze = await executeFreezeTool(this.host, call);
          if (freeze) return freeze;
          const pageLayout = await executePageLayoutTool(this.host, call);
          if (pageLayout) return pageLayout;
          const shape = await executeShapeTool(this.host, call);
          if (shape) return shape;
          const tableUnlist = await executeTableUnlistTool(this.host, call);
          if (tableUnlist) return tableUnlist;
          return fail(call.name, `Unknown tool: ${String((call as ToolCall).name)}`);
        }
      }
    } catch (error) {
      return fail(call.name, error instanceof Error ? error.message : String(error));
    }
  }
}
