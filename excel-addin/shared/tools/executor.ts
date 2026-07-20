import type { HostAdapter, RangeExpandMode, RangeFormat } from "../host/types";
import type { CellValue, ToolCall, ToolFailure, ToolName, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";
import {
  optionalIdent,
  rejectUnknownCoreToolArguments,
  rejectUnknownRangeFormatFields,
  requireIdent,
  requireValueString,
} from "./argValidation";
import { executeSheetOperation } from "./sheetOperation";
import { executeValidationTool } from "./validationExecutor";
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
import { executeChartSeriesTrendlineTool } from "./chartSeriesTrendlineExecutor";
import { executeChartSeriesTrendlineFormatTool } from "./chartSeriesTrendlineFormatExecutor";
import { executeChartSeriesMarkersTool } from "./chartSeriesMarkersExecutor";
import { executeChartImageTool } from "./chartImageExecutor";
import { executeRangeImageTool } from "./rangeImageExecutor";
import { executeRangeStructureTool } from "./rangeStructureExecutor";
import { executeChartSourceTool } from "./chartSourceExecutor";
import { executeObjectUpdateTool } from "./objectUpdateExecutor";
import { executePageLayoutTool } from "./pageLayoutExecutor";
import { executeShapeTool } from "./shapeExecutor";
import { executeStructureTool } from "./structureExecutor";
import { executeTableUnlistTool } from "./tableUnlistExecutor";
import { executeTableFilterTool } from "./tableFilterExecutor";
import { executeTableSortTool } from "./tableSortExecutor";
import { executeFormulaProtectionTool } from "./formulaProtectionExecutor";
import { executeFormulaGovernanceTool } from "./formulaGovernanceExecutor";
import { executePivotTool } from "./pivotExecutor";
import { writeFormulaWithVerify, writeRangeWithVerify } from "./writeWithVerify";

function fail(tool: ToolName, error: string, detail?: unknown): ToolFailure {
  return { ok: false, tool, error, detail };
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
  const raw = args.expand;
  if (typeof raw !== "string") {
    throw new Error("expand must be none|spill|currentArray|currentRegion");
  }
  const value = raw.trim();
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
  const raw = args.operation;
  if (typeof raw !== "string") {
    throw new Error("operation must be add|rename|delete|copy|move");
  }
  const value = raw.trim();
  if (
    value !== "add" &&
    value !== "rename" &&
    value !== "delete" &&
    value !== "copy" &&
    value !== "move"
  ) {
    throw new Error("operation must be add|rename|delete|copy|move");
  }
  return value as "add" | "rename" | "delete" | "copy" | "move";
}

function fromHost(
  tool: Parameters<typeof mapHostResultToToolResult>[0],
  result: Parameters<typeof mapHostResultToToolResult>[1],
): ReturnType<typeof mapHostResultToToolResult> {
  return mapHostResultToToolResult(tool, result);
}

function parseMaxItemsPerCategory(args: Record<string, unknown>): number {
  if (!Object.prototype.hasOwnProperty.call(args, "maxItemsPerCategory") || args.maxItemsPerCategory === undefined) {
    return 100;
  }
  const value = args.maxItemsPerCategory;
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1 || value > 500) {
    throw new Error("maxItemsPerCategory must be an integer between 1 and 500");
  }
  return value;
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
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "range"),
              optionalExpand(call.arguments),
            ),
          );
        case "range.write":
          return await writeRangeWithVerify(this.host, {
            sheetName: requireIdent(call.arguments, "sheetName"),
            range: requireIdent(call.arguments, "range"),
            values: requireValues(call.arguments),
            verify: call.arguments.verify !== false,
          });
        case "range.clear":
          return fromHost(
            call.name,
            await this.host.clearRange(
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "range"),
            ),
          );
        case "range.format.read":
          return fromHost(
            call.name,
            await this.host.readFormat(
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "range"),
            ),
          );
        case "range.format.write":
          return fromHost(
            call.name,
            await this.host.writeFormat(
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "range"),
              requireFormat(call.arguments),
            ),
          );
        case "formula.read": {
          const read = await this.host.readRange(
            requireIdent(call.arguments, "sheetName"),
            requireIdent(call.arguments, "range"),
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
            sheetName: requireIdent(call.arguments, "sheetName"),
            range: requireIdent(call.arguments, "range"),
            formula: requireValueString(call.arguments, "formula"),
            verify: call.arguments.verify !== false,
          });
        case "formula.context":
          return fromHost(
            call.name,
            await this.host.getFormulaContext(
              requireIdent(call.arguments, "sheetName"),
              optionalIdent(call.arguments, "range"),
            ),
          );
        case "sheet.list":
          return fromHost(call.name, await this.host.listSheets());
        case "sheet.add":
          return fromHost(
            call.name,
            await this.host.addSheet(requireIdent(call.arguments, "sheetName")),
          );
        case "sheet.rename":
          return fromHost(
            call.name,
            await this.host.renameSheet(
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "newName"),
            ),
          );
        case "sheet.delete":
          return fromHost(
            call.name,
            await this.host.deleteSheet(requireIdent(call.arguments, "sheetName")),
          );
        case "sheet.operation":
          return await executeSheetOperation(this.host, call.arguments, fromHost, {
            requireString: requireIdent,
            optionalString: optionalIdent,
            optionalFiniteNumber,
            requireSheetOperation,
          });
        case "table.list":
          return fromHost(
            call.name,
            await this.host.listTables(optionalIdent(call.arguments, "sheetName")),
          );
        case "table.create":
          return fromHost(
            call.name,
            await this.host.createTable({
              sheetName: requireIdent(call.arguments, "sheetName"),
              address: requireIdent(call.arguments, "range"),
              name: optionalIdent(call.arguments, "name"),
              hasHeaders: optionalBoolean(call.arguments, "hasHeaders"),
            }),
          );
        case "table.delete":
          return fromHost(
            call.name,
            await this.host.deleteTable(
              requireIdent(call.arguments, "sheetName"),
              requireIdent(call.arguments, "tableName"),
            ),
          );
        case "workbook.inspect":
          return fromHost(call.name, await this.host.inspectWorkbook());
        case "workbook.objects.inspect":
          return fromHost(
            call.name,
            await this.host.inspectWorkbookObjects({
              maxItemsPerCategory: parseMaxItemsPerCategory(call.arguments),
              sheetName: optionalIdent(call.arguments, "sheetName"),
            }),
          );
        case "workbook.save":
          return fromHost(call.name, await this.host.saveWorkbook());
        default: {
          const validation = await executeValidationTool(this.host, call);
          if (validation) return validation;
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
          const chartSeriesTrendlines = await executeChartSeriesTrendlineTool(this.host, call);
          if (chartSeriesTrendlines) return chartSeriesTrendlines;
          const chartSeriesTrendlineFormat = await executeChartSeriesTrendlineFormatTool(this.host, call);
          if (chartSeriesTrendlineFormat) return chartSeriesTrendlineFormat;
          const chartSeriesMarkers = await executeChartSeriesMarkersTool(this.host, call);
          if (chartSeriesMarkers) return chartSeriesMarkers;
          const chartImage = await executeChartImageTool(this.host, call);
          if (chartImage) return chartImage;
          const rangeImage = await executeRangeImageTool(this.host, call);
          if (rangeImage) return rangeImage;
          const rangeStructure = await executeRangeStructureTool(this.host, call);
          if (rangeStructure) return rangeStructure;
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
          const tableFilter = await executeTableFilterTool(this.host, call);
          if (tableFilter) return tableFilter;
          const tableSort = await executeTableSortTool(this.host, call);
          if (tableSort) return tableSort;
          const formulaProtection = await executeFormulaProtectionTool(this.host, call);
          if (formulaProtection) return formulaProtection;
          const formulaGovernance = await executeFormulaGovernanceTool(this.host, call);
          if (formulaGovernance) return formulaGovernance;
          const pivot = await executePivotTool(this.host, call);
          if (pivot) return pivot;
          return fail(call.name, `Unknown tool: ${String((call as ToolCall).name)}`);
        }
      }
    } catch (error) {
      return fail(call.name, error instanceof Error ? error.message : String(error));
    }
  }
}
