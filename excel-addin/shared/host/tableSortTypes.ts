/** Table sort types (Office.js Excel.Table.sort, ExcelApi 1.2+). */

export interface TableSortFieldInput {
  /** 1-based column index within the table (header left = 1). */
  columnIndex: number;
  ascending?: boolean;
}

export interface TableSortApplyInput {
  sheetName: string;
  tableName: string;
  fields: TableSortFieldInput[];
  matchCase?: boolean;
}

export interface TableSortClearInput {
  sheetName: string;
  tableName: string;
}

export interface TableSortGetInput {
  sheetName: string;
  tableName: string;
}

export interface TableSortFieldInfo {
  columnIndex: number;
  ascending: boolean;
}

export interface TableSortInfo {
  sheetName: string;
  tableName: string;
  fields: TableSortFieldInfo[];
}
