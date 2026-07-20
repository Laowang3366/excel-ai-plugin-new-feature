import {
  type CellValue,
  type HostResult,
  ok,
  unsupported,
} from "./types";

/**
 * WPS JSA surface helpers.
 * Verified in-repo: Application / ActiveWorkbook / Name / JSIDE CodeModule (desktop bridge).
 * Range Value2/Formula/Clear, UsedRange, Worksheets, and optional ET COM members below are
 * assumed with runtime member probes (not device-verified).
 */
export interface WpsFont {
  Name?: string | null;
  Size?: number | null;
  Bold?: boolean | null;
  Color?: number | null;
}

export interface WpsInterior {
  Color?: number | null;
}

export interface WpsRangeCollection {
  AutoFit?: () => void;
}

export interface WpsRange {
  Address?: string;
  Value2?: unknown;
  Formula?: unknown;
  Clear?: () => void;
  /** Excel/WPS COM CurrentRegion (surrounding contiguous region). */
  CurrentRegion?: WpsRange;
  Font?: WpsFont;
  Interior?: WpsInterior;
  NumberFormat?: string | unknown;
  HorizontalAlignment?: number | string | null;
  VerticalAlignment?: number | string | null;
  WrapText?: boolean | null;
  ColumnWidth?: number | null;
  RowHeight?: number | null;
  Columns?: WpsRangeCollection;
  Rows?: WpsRangeCollection;
  EntireColumn?: WpsRangeCollection;
  EntireRow?: WpsRangeCollection;
  /** Excel/WPS COM Range.Insert(Shift). */
  Insert?: (shift?: number | string) => WpsRange | void;
  /** Excel/WPS COM Range.Delete(Shift). */
  Delete?: (shift?: number | string) => void;
}

export interface WpsSheet {
  Name: string;
  Index?: number;
  Range: (address: string) => WpsRange;
  Delete?: () => void;
  UsedRange?: WpsRange;
  /** Excel/WPS COM: Copy(Before?, After?). */
  Copy?: (before?: WpsSheet | undefined, after?: WpsSheet | undefined) => void;
  /** Excel/WPS COM: Move(Before?, After?). */
  Move?: (before?: WpsSheet | undefined, after?: WpsSheet | undefined) => void;
  /** xlSheetVisible=-1, Hidden=0, VeryHidden=2 (desktop ExcelObjectActionService). */
  Visible?: number | string;
  /** True when sheet contents are protected. */
  ProtectContents?: boolean;
  Protect?: (password?: string, ...rest: unknown[]) => void;
  Unprotect?: (password?: string) => void;
  Names?: WpsNames;
}

export interface WpsSheets {
  Count: number;
  Item: (indexOrName: number | string) => WpsSheet;
  Add?: (before?: unknown, after?: unknown, count?: number, type?: unknown) => WpsSheet;
}

export interface WpsName {
  Name?: string;
  RefersTo?: string;
  Visible?: boolean;
  Delete?: () => void;
}

export interface WpsNames {
  Count?: number;
  Item?: (indexOrName: number | string) => WpsName;
  Add?: (name: string, refersTo?: string, ...rest: unknown[]) => WpsName;
}

export interface WpsWorkbook {
  Name: string;
  ActiveSheet: WpsSheet;
  Worksheets: WpsSheets;
  Names?: WpsNames;
}

export interface WpsApplication {
  ActiveWorkbook?: WpsWorkbook;
  Selection?: WpsRange & { Worksheet?: WpsSheet };
  Name?: string;
}

declare global {
  interface Window {
    Application?: WpsApplication;
  }
}

export function getApplication(): WpsApplication | null {
  if (typeof window === "undefined") return null;
  return window.Application ?? null;
}

export function matrixFrom(value: unknown): CellValue[][] {
  if (value == null) return [[null]];
  if (!Array.isArray(value)) return [[value as CellValue]];
  if (value.length === 0) return [];
  if (!Array.isArray(value[0])) return [value as CellValue[]];
  return (value as unknown[][]).map((row) =>
    row.map((cell) => (cell === undefined ? null : (cell as CellValue))),
  );
}

export function formulaMatrixFrom(value: unknown): string[][] {
  return matrixFrom(value).map((row) => row.map((cell) => String(cell ?? "")));
}

export function requireApp(capability: string): HostResult<WpsApplication> {
  const app = getApplication();
  if (!app) {
    return unsupported(
      capability,
      "wps-jsa",
      "window.Application is not available",
      "desktop/public/wps-jsa-bridge uses Application; in-process JSA task pane must expose it",
    );
  }
  return ok(app);
}

export function requireWorkbook(capability: string): HostResult<WpsWorkbook> {
  const appResult = requireApp(capability);
  if (!appResult.ok) return appResult;
  const workbook = appResult.data.ActiveWorkbook;
  if (!workbook) {
    return unsupported(capability, "wps-jsa", "ActiveWorkbook is missing", "WPS JSA Application");
  }
  return ok(workbook);
}

export function getSheet(workbook: WpsWorkbook, sheetName: string): WpsSheet | null {
  try {
    return workbook.Worksheets.Item(sheetName);
  } catch {
    return null;
  }
}
