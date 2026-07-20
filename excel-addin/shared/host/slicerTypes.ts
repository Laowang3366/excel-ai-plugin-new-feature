/**
 * Public slicer DTOs (ExcelApi 1.10 official stable subset).
 * No sourceName/sourceType/sourceField host readback — stable Slicer has none.
 */

export type SlicerSourceType = "table" | "pivotTable";

/** Public camelCase; host tokens DataSourceOrder|Ascending|Descending. */
export type SlicerSortBy = "dataSourceOrder" | "ascending" | "descending";

export const SLICER_SORT_BY_VALUES = [
  "dataSourceOrder",
  "ascending",
  "descending",
] as const;

export const SLICER_MAX_NAME_LEN = 255;
export const SLICER_MAX_CAPTION_LEN = 255;
export const SLICER_MAX_STYLE_LEN = 255;
export const SLICER_MAX_FILTER_KEYS = 500;
export const SLICER_MAX_ITEMS_READBACK = 1000;

export type SlicerRequestedSource = {
  sourceType: SlicerSourceType;
  sourceName: string;
  sourceField: string;
};

export type SlicerInfo = {
  name: string;
  id: string;
  caption: string;
  sheetName: string;
  top: number;
  left: number;
  width: number;
  height: number;
  sortBy: SlicerSortBy;
  style: string;
  isFilterCleared: boolean;
  selectedItemCount?: number;
  itemCount?: number;
  limitations?: string[];
};

export type SlicerListInput = {
  sheetName?: string;
};

export type SlicerListInfo = {
  slicers: SlicerInfo[];
  limitations?: string[];
};

export type SlicerCreateInput = {
  /** Required for interactive slicer create (desktop parity). */
  advancedIntent: "interactive-pivot";
  sourceType: SlicerSourceType;
  sourceName: string;
  /** Table column name or PivotField name (string only; no raw index). */
  sourceField: string;
  destinationSheet: string;
  name?: string;
  caption?: string;
  top?: number;
  left?: number;
  width?: number;
  height?: number;
  style?: string;
  sortBy?: SlicerSortBy;
};

export type SlicerCreateInfo = SlicerInfo & {
  requestedSource: SlicerRequestedSource;
  verification: {
    ok: boolean;
    objectExists: boolean;
    nameMatches: boolean;
    checks: Array<{ name: string; ok: boolean; message?: string }>;
  };
};

export type SlicerUpdateInput = {
  name: string;
  newName?: string;
  caption?: string;
  top?: number;
  left?: number;
  width?: number;
  height?: number;
  style?: string;
  sortBy?: SlicerSortBy;
};

export type SlicerDeleteInput = {
  name: string;
};

export type SlicerDeleteInfo = {
  deleted: string;
};

export type SlicerItemInfo = {
  key: string;
  name: string;
  isSelected: boolean;
  hasData: boolean;
};

export type SlicerFilterGetInput = {
  name: string;
};

export type SlicerFilterInfo = {
  name: string;
  isFilterCleared: boolean;
  /** Official getSelectedItems returns keys. */
  selectedKeys: string[];
  items: SlicerItemInfo[];
  itemCount: number;
  truncated: boolean;
  verified: boolean;
  limitations?: string[];
};

export type SlicerFilterApplyInput = {
  name: string;
  /**
   * Item keys for Slicer.selectItems.
   * Empty array = select all (official Office.js semantics), not "select none".
   */
  keys: string[];
};

export type SlicerFilterClearInput = {
  name: string;
};
