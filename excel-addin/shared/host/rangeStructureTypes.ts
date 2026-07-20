export type RangeInsertShift = "down" | "right";
export type RangeDeleteShift = "up" | "left";
export type RangeAutofitDirection = "rows" | "columns" | "both";

export interface RangeInsertInput {
  sheetName: string;
  address: string;
  shift: RangeInsertShift;
}

export interface RangeDeleteInput {
  sheetName: string;
  address: string;
  shift: RangeDeleteShift;
}

export interface RangeMutationInfo {
  sheetName: string;
  address: string;
  shift: RangeInsertShift | RangeDeleteShift;
  operation: "insert" | "delete";
}

export interface RangeAutofitInput {
  sheetName: string;
  address: string;
  direction: RangeAutofitDirection;
}

export interface RangeAutofitInfo {
  sheetName: string;
  address: string;
  direction: RangeAutofitDirection;
  columnWidth: number | null;
  rowHeight: number | null;
}
