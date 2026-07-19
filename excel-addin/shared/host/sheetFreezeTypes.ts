export type SheetFreezeCommand = "rows" | "columns" | "at" | "clear";

export interface SheetFreezeInfo {
  sheetName: string;
  address: string | null;
  rowCount: number;
  columnCount: number;
}

export interface SheetFreezeSetInput {
  sheetName: string;
  command: SheetFreezeCommand;
  count?: number;
  address?: string;
}
