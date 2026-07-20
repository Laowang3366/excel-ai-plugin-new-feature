/**
 * Minimal Office.js PivotTable surface used by officeJsPivot* (ExcelApi 1.8).
 * Not a full Excel.d.ts; only members we queue/load.
 */
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";

export type ExcelAggregationFunction =
  | "Sum"
  | "Count"
  | "Average"
  | "Max"
  | "Min"
  | "Product"
  | "CountNumbers"
  | "StdDev"
  | "StdDevP"
  | "Var"
  | "VarP"
  | "Unknown"
  | "Automatic";

export interface ExcelPivotHierarchy {
  name: string;
  load(props: string): void;
}

export interface ExcelPivotHierarchyCollection {
  items: ExcelPivotHierarchy[];
  count?: number;
  load(props: string): void;
  getItem(name: string): ExcelPivotHierarchy;
  add(hierarchy: ExcelPivotHierarchy): ExcelPivotHierarchy | void;
}

export interface ExcelDataPivotHierarchy {
  name: string;
  summarizeBy: ExcelAggregationFunction | string;
  load(props: string): void;
}

export interface ExcelDataPivotHierarchyCollection {
  items: ExcelDataPivotHierarchy[];
  count?: number;
  load(props: string): void;
  add(hierarchy: ExcelPivotHierarchy): ExcelDataPivotHierarchy;
}

export interface ExcelPivotLayout {
  getRange(): ExcelRange;
}

export interface ExcelPivotTable {
  name: string;
  id?: string;
  worksheet?: { name: string; load(props: string): void };
  hierarchies: ExcelPivotHierarchyCollection;
  rowHierarchies: ExcelPivotHierarchyCollection;
  columnHierarchies: ExcelPivotHierarchyCollection;
  filterHierarchies: ExcelPivotHierarchyCollection;
  dataHierarchies: ExcelDataPivotHierarchyCollection;
  layout: ExcelPivotLayout;
  /** Optional; when missing, source is reported as limitation. */
  getDataSourceString?: () => string;
  refresh(): void;
  load(props: string): void;
}

export interface ExcelPivotTableCollection {
  items: ExcelPivotTable[];
  count?: number;
  load(props: string): void;
  getItem(name: string): ExcelPivotTable;
  add(name: string, source: ExcelRange, destination: ExcelRange): ExcelPivotTable;
  refreshAll(): void;
}

export type ExcelWorksheetWithPivot = ExcelWorksheet & {
  pivotTables: ExcelPivotTableCollection;
};

export type ExcelDataConnectionCollection = {
  /** ExcelApi 1.7 — no official items/count readback contract. */
  refreshAll(): void;
};

export type ExcelRequestContextWithPivot = ExcelRequestContext & {
  workbook: ExcelRequestContext["workbook"] & {
    /** ExcelApi 1.7 DataConnectionCollection (optional member). */
    dataConnections?: ExcelDataConnectionCollection;
    worksheets: ExcelRequestContext["workbook"]["worksheets"] & {
      getItem(name: string): ExcelWorksheetWithPivot;
      getItemOrNullObject?(name: string): ExcelWorksheetWithPivot & { isNullObject?: boolean };
      add(name?: string): ExcelWorksheetWithPivot;
    };
  };
};

export type { ExcelRange, ExcelRequestContext, ExcelWorksheet };
