/** Table AutoFilter types (Office.js Excel.Table.autoFilter, ExcelApi 1.2+). */

/** Public filter modes we implement; color/icon/dynamic remain typed unsupported. */
export const TABLE_FILTER_ON = [
  "values",
  "custom",
  "topItems",
  "bottomItems",
  "topPercent",
  "bottomPercent",
] as const;

export type TableFilterOn = (typeof TABLE_FILTER_ON)[number];

export function isTableFilterOn(value: string): value is TableFilterOn {
  return (TABLE_FILTER_ON as readonly string[]).includes(value);
}

export type TableFilterOperator = "and" | "or";

export interface TableFilterApplyInput {
  sheetName: string;
  tableName: string;
  /** 1-based column index within the table (header left = 1). */
  columnIndex: number;
  filterOn: TableFilterOn;
  /** Required when filterOn is values. */
  values?: string[];
  /** Custom criteria (Excel criterion1 / criterion2 strings). */
  criterion1?: string;
  criterion2?: string;
  operator?: TableFilterOperator;
  /** Positive count/percent for top/bottom modes. */
  threshold?: number;
}

export interface TableFilterClearInput {
  sheetName: string;
  tableName: string;
}

export interface TableFilterGetInput {
  sheetName: string;
  tableName: string;
}

export interface TableFilterInfo {
  sheetName: string;
  tableName: string;
  /** Host AutoFilter.enabled (ExcelApi 1.9). */
  enabled: boolean;
  /** Echo of last applied public columnIndex when apply just ran; get may omit. */
  columnIndex?: number;
  filterOn?: TableFilterOn;
}
