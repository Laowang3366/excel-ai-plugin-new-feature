import type {
  CellValue,
  ChartType,
  HostResult,
  RangeData,
  RangeFormat,
  SelectionInfo,
  SheetInfo,
} from "../host/types";

export type ToolName =
  | "host.status"
  | "selection.get"
  | "range.read"
  | "range.write"
  | "range.clear"
  | "range.insert"
  | "range.delete"
  | "range.autofit"
  | "range.format.read"
  | "range.format.write"
  | "formula.read"
  | "formula.write"
  | "formula.context"
  | "formula.protection.inspect"
  | "formula.protection.manage"
  | "formula.dependencies.inspect"
  | "formula.references.repair"
  | "formula.convertToValues"
  | "formula.backups.inspect"
  | "formula.backups.restore"
  | "sheet.list"
  | "sheet.add"
  | "sheet.rename"
  | "sheet.delete"
  | "sheet.operation"
  | "table.list"
  | "table.create"
  | "table.delete"
  | "table.unlist"
  | "table.filter.get"
  | "table.filter.apply"
  | "table.filter.clear"
  | "table.sort.get"
  | "table.sort.apply"
  | "table.sort.clear"
  | "chart.list"
  | "chart.create"
  | "chart.delete"
  | "chart.series.list"
  | "chart.series.update"
  | "chart.source.update"
  | "chart.axes.update"
  | "chart.series.dataLabels.update"
  | "chart.series.axisGroup.update"
  | "chart.series.delete"
  | "chart.series.add"
  | "chart.series.values.update"
  | "chart.series.bubbleSizes.update"
  | "chart.image.get"
  | "range.image.get"
  | "workbook.inspect"
  | "workbook.objects.inspect"
  | "workbook.save"
  | "conditionalFormat.list"
  | "conditionalFormat.add"
  | "conditionalFormat.delete"
  | "dataValidation.read"
  | "dataValidation.write"
  | "dataValidation.clear"
  | "sheet.visibility.get"
  | "sheet.visibility.set"
  | "sheet.protection.get"
  | "sheet.protection.protect"
  | "sheet.protection.unprotect"
  | "namedRange.list"
  | "namedRange.create"
  | "namedRange.update"
  | "namedRange.delete"
  | "table.update"
  | "chart.update"
  | "sheet.display.get"
  | "sheet.display.set"
  | "sheet.freeze.get"
  | "sheet.freeze.set"
  | "sheet.pageLayout.get"
  | "sheet.pageLayout.set"
  | "shape.list"
  | "shape.create"
  | "shape.delete"
  | "shape.update";

export type RiskLevel = "safe" | "moderate" | "dangerous";

export interface ToolDefinition {
  name: ToolName;
  description: string;
  riskLevel: RiskLevel;
  parameters: Record<string, unknown>;
}

export interface ToolCall {
  name: ToolName;
  arguments: Record<string, unknown>;
}

export interface ToolSuccess<T = unknown> {
  ok: true;
  tool: ToolName;
  data: T;
  verification?: unknown;
}

export interface ToolFailure {
  ok: false;
  tool: ToolName;
  unsupported?: true;
  error: string;
  detail?: unknown;
}

export type ToolResult = ToolSuccess | ToolFailure;

export interface RangeWriteArgs {
  sheetName: string;
  range: string;
  values: CellValue[][];
  verify?: boolean;
}

export interface FormulaWriteArgs {
  sheetName: string;
  range: string;
  formula: string;
  verify?: boolean;
}

export type {
  CellValue,
  ChartType,
  HostResult,
  RangeData,
  RangeFormat,
  SelectionInfo,
  SheetInfo,
};
