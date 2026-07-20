/**
 * Minimal Office.js Slicer surface (ExcelApi 1.10). Not full Excel.d.ts.
 */
import type { ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";

export type ExcelSlicerSortHost = "DataSourceOrder" | "Ascending" | "Descending" | string;

export interface ExcelSlicerItem {
  name: string;
  key: string;
  isSelected: boolean;
  hasData: boolean;
  load(props: string): void;
}

export interface ExcelSlicerItemCollection {
  items: ExcelSlicerItem[];
  load(props: string): void;
  getItem?(key: string): ExcelSlicerItem;
}

export interface ExcelSlicer {
  name: string;
  id: string;
  caption: string;
  height: number;
  width: number;
  top: number;
  left: number;
  sortBy: ExcelSlicerSortHost;
  style: string;
  isFilterCleared: boolean;
  worksheet: { name: string; load(props: string): void };
  slicerItems: ExcelSlicerItemCollection;
  load(props: string): void;
  delete(): void;
  clearFilters(): void;
  getSelectedItems(): { value: string[] }; // value only after context.sync (ClientResult)
  selectItems(items?: string[]): void;
}

export interface ExcelSlicerCollection {
  items: ExcelSlicer[];
  load(props: string): void;
  add(
    slicerSource: string | object,
    sourceField: string | number | object,
    slicerDestination?: string | object,
  ): ExcelSlicer;
  getItem(key: string): ExcelSlicer;
  getItemOrNullObject?(key: string): ExcelSlicer & { isNullObject?: boolean };
}

export type ExcelTableLike = {
  name?: string;
  columns?: {
    getItem(name: string): object;
    getItemOrNullObject?(name: string): object & { isNullObject?: boolean };
  };
  load?(props: string): void;
};

export type ExcelPivotLike = {
  name?: string;
  hierarchies?: { getItem(name: string): object };
  load?(props: string): void;
};

export type ExcelWorksheetWithSlicers = ExcelWorksheet & {
  slicers: ExcelSlicerCollection;
  tables?: {
    getItem(name: string): ExcelTableLike;
    getItemOrNullObject?(name: string): ExcelTableLike & { isNullObject?: boolean };
  };
  pivotTables?: {
    getItem(name: string): ExcelPivotLike;
    getItemOrNullObject?(name: string): ExcelPivotLike & { isNullObject?: boolean };
  };
};

export type ExcelRequestContextWithSlicer = ExcelRequestContext & {
  workbook: ExcelRequestContext["workbook"] & {
    slicers: ExcelSlicerCollection;
    tables?: {
      getItem(name: string): ExcelTableLike;
      getItemOrNullObject?(name: string): ExcelTableLike & { isNullObject?: boolean };
    };
    pivotTables?: {
      getItem(name: string): ExcelPivotLike;
      getItemOrNullObject?(name: string): ExcelPivotLike & { isNullObject?: boolean };
    };
    worksheets: ExcelRequestContext["workbook"]["worksheets"] & {
      getItem(name: string): ExcelWorksheetWithSlicers;
      items?: ExcelWorksheetWithSlicers[];
      load(props: string): void;
    };
  };
};

export type { ExcelRequestContext, ExcelWorksheet };
