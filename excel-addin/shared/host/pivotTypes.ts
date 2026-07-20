/**
 * Pivot table lifecycle contracts for the Excel add-in (current workbook only).
 * Office.js ExcelApi 1.8: WorksheetPivotTableCollection.add + hierarchy layout.
 */

export type PivotAggregationFunction = "sum" | "count" | "average" | "max" | "min";

export type PivotFieldObject = {
  name: string;
  function?: PivotAggregationFunction;
  caption?: string;
};

/** String field name or object (dataFields may set function/caption). */
export type PivotFieldSpec = string | PivotFieldObject;

export type PivotNormalizedField = {
  name: string;
  function?: PivotAggregationFunction;
  caption?: string;
};

export type PivotHierarchySummary = {
  name: string;
  summarizeBy?: string;
  caption?: string;
};

export type PivotTableInfo = {
  name: string;
  sheetName: string;
  source?: string | null;
  destination?: string | null;
  rowFields: PivotHierarchySummary[];
  columnFields: PivotHierarchySummary[];
  filterFields: PivotHierarchySummary[];
  dataFields: PivotHierarchySummary[];
  /** True after a successful refresh call in this session when known. */
  refreshed?: boolean | null;
  limitations?: string[];
};

export type PivotListInput = {
  sheetName?: string;
};

export type PivotListInfo = {
  pivots: PivotTableInfo[];
  limitations: string[];
};

export type PivotCreateInput = {
  sourceSheetName: string;
  sourceAddress: string;
  name?: string;
  /** Bare A1 on source sheet, or Sheet!A1 / 'Sheet'!A1. Empty → Pivots sheet auto address. */
  destination?: string;
  rowFields?: PivotFieldSpec[];
  columnFields?: PivotFieldSpec[];
  filterFields?: PivotFieldSpec[];
  dataFields?: PivotFieldSpec[];
  /** Desktop parity: must be "interactive-pivot" when provided by tools. */
  advancedIntent?: "interactive-pivot";
};

export type PivotCreateVerification = {
  ok: boolean;
  objectExists: boolean;
  nameMatches: boolean;
  destinationReadable: boolean;
  rowFieldCount: number;
  columnFieldCount: number;
  filterFieldCount: number;
  dataFieldCount: number;
  checks: Array<{ name: string; ok: boolean; message?: string }>;
};

export type PivotCreateInfo = {
  name: string;
  sheetName: string;
  source: string;
  destination: string;
  rowFields: PivotHierarchySummary[];
  columnFields: PivotHierarchySummary[];
  filterFields: PivotHierarchySummary[];
  dataFields: PivotHierarchySummary[];
  verification: PivotCreateVerification;
  limitations?: string[];
};

export type PivotRefreshInput = {
  sheetName?: string;
  name?: string;
  /**
   * Desktop maps true → Workbook.RefreshAll. Office.js has no proven equivalent;
   * true is rejected at the tool/host boundary.
   */
  refreshConnections?: boolean;
  advancedIntent?: "interactive-pivot";
};

export type PivotRefreshEntry = {
  name: string;
  sheetName: string;
  refreshed: boolean;
};

export type PivotRefreshInfo = {
  refreshed: PivotRefreshEntry[];
  count: number;
  limitations?: string[];
};

export const PIVOT_AGGREGATION_FUNCTIONS: readonly PivotAggregationFunction[] = [
  "sum",
  "count",
  "average",
  "max",
  "min",
] as const;

export const PIVOT_DEFAULT_SHEET = "Pivots";
export const PIVOT_MAX_FIELDS = 64;
