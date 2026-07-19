import {
  type CellValue,
  type HostResult,
  type RangeFormat,
  ok,
  unsupported,
} from "./types";

export type ExcelRunFn = <T>(batch: (context: ExcelRequestContext) => Promise<T>) => Promise<T>;

export interface ExcelFont {
  name: string;
  size: number;
  bold: boolean;
  color: string;
  load(props: string): void;
}

export interface ExcelFill {
  color: string;
  load(props: string): void;
}

export interface ExcelRangeFormat {
  font: ExcelFont;
  fill: ExcelFill;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  load(props: string): void;
}

export interface ExcelConditionalFormat {
  id: string;
  type: string;
  cellValue?: {
    rule: { formula1: string; formula2?: string; operator: string };
    format: {
      fill: { color: string };
      font: { color: string };
    };
  };
  /** Office.js: ConditionalFormatRule.formula is a string (not nested .formula.formula). */
  custom?: {
    rule: { formula: string };
    format: {
      fill: { color: string };
      font: { color: string };
    };
  };
  getRange(): ExcelRange;
  delete(): void;
  load(props: string): void;
}

/** Official DataValidation surface: type/rule/ignoreBlanks/errorAlert/prompt — no top-level formula1. */
export interface ExcelDataValidationRule {
  list?: { inCellDropDown?: boolean; source?: string | unknown };
  wholeNumber?: {
    formula1?: string | number;
    formula2?: string | number;
    operator?: string;
  };
}

export interface ExcelDataValidation {
  type: string | null;
  ignoreBlanks: boolean;
  rule: ExcelDataValidationRule;
  errorAlert?: { message?: string; showAlert?: boolean; style?: string; title?: string };
  prompt?: { message?: string; showPrompt?: boolean; title?: string };
  load(props: string): void;
  clear(): void;
}

export interface ExcelRange {
  address: string;
  values: CellValue[][];
  formulas: string[][];
  numberFormat: string[][] | string;
  rowCount: number;
  columnCount: number;
  format: ExcelRangeFormat;
  conditionalFormats: {
    items: ExcelConditionalFormat[];
    add(type: string): ExcelConditionalFormat;
    getItem(id: string): ExcelConditionalFormat;
    getItemAt(index: number): ExcelConditionalFormat;
    getCount(): { value: number; load(props: string): void };
    clearAll(): void;
    load(props: string): void;
  };
  dataValidation: ExcelDataValidation;
  load(props: string): void;
  clear(): void;
  getSpillingToRange(): ExcelRange;
  getSurroundingRegion(): ExcelRange;
  getCurrentArray(): ExcelRange;
  getCell(row: number, column: number): ExcelRange;
}

export interface ExcelTable {
  name: string;
  showHeaders: boolean;
  showFilterButton: boolean;
  showTotals: boolean;
  style: string;
  getRange(): ExcelRange;
  delete(): void;
  /** ExcelApi 1.2: convert table to plain range; keeps cell values. */
  convertToRange(): ExcelRange;
  load(props: string): void;
}

export interface ExcelChart {
  name: string;
  chartType: string;
  style: number;
  title: { text: string; visible?: boolean; load(props: string): void };
  legend: { visible: boolean; load(props: string): void };
  left: number;
  top: number;
  width: number;
  height: number;
  delete(): void;
  load(props: string): void;
}

/** Official: NamedItem.name is readonly; formula/visible are writable. */
export interface ExcelNamedItem {
  readonly name: string;
  formula: string;
  visible: boolean;
  load(props: string): void;
  delete(): void;
}

export interface ExcelNamedItemCollection {
  items: ExcelNamedItem[];
  add(name: string, reference: string | ExcelRange): ExcelNamedItem;
  getItem(name: string): ExcelNamedItem;
  load(props: string): void;
}

/** Official: protect(options?: WorksheetProtectionOptions, password?: string). */
export interface ExcelWorksheetProtection {
  protected: boolean;
  load(props: string): void;
  protect(options?: object, password?: string): void;
  unprotect(password?: string): void;
}

/** Official Worksheet.freezePanes. */
export interface ExcelFreezePanes {
  freezeRows(count: number): void;
  freezeColumns(count: number): void;
  freezeAt(range: ExcelRange): void;
  unfreeze(): void;
  getLocationOrNullObject(): ExcelRange & { isNullObject: boolean };
}

/** Official PageLayoutZoomOptions (scale may be null under fit-to-pages). */
export interface ExcelPageLayoutZoomOptions {
  scale?: number | null;
  horizontalFitToPages?: number;
  verticalFitToPages?: number;
}

/** Official RangeAreas facade for getPrintAreaOrNullObject only. */
export interface ExcelRangeAreas {
  address: string;
  isNullObject: boolean;
  load(props: string): void;
}


/** ExcelApi 1.9 HeaderFooter (default page slots only). */
export interface ExcelHeaderFooter {
  leftHeader: string;
  centerHeader: string;
  rightHeader: string;
  leftFooter: string;
  centerFooter: string;
  rightFooter: string;
  load(props: string): void;
}

export interface ExcelHeaderFooterGroup {
  defaultForAllPages: ExcelHeaderFooter;
  load(props: string): void;
}

export interface ExcelPageLayout {
  orientation: string;
  centerHorizontally: boolean;
  centerVertically: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  blackAndWhite: boolean;
  /** ExcelApi 1.9 draft quality print. */
  draftMode: boolean;
  /** ExcelApi 1.9 PrintOrder: DownThenOver | OverThenDown. */
  printOrder: string;
  /** ExcelApi 1.9; host may return "" for auto. */
  firstPageNumber: number | string | null;
  topMargin: number;
  bottomMargin: number;
  leftMargin: number;
  rightMargin: number;
  headerMargin: number;
  footerMargin: number;
  /** ExcelApi 1.9 PaperType string (e.g. A4, Letter). */
  paperSize: string;
  /** ExcelApi 1.9 headers/footers group (defaultForAllPages only). */
  headersFooters: ExcelHeaderFooterGroup;
  /** Assign whole object: pageLayout.zoom = { scale } or { horizontalFitToPages, verticalFitToPages }. Never write zoom.* subprops. */
  zoom: ExcelPageLayoutZoomOptions;
  load(props: string): void;
  /** Official: returns RangeAreas. */
  getPrintAreaOrNullObject(): ExcelRangeAreas;
  setPrintArea(address: string): void;
  /** Official: returns Range (not RangeAreas). */
  getPrintTitleRowsOrNullObject(): ExcelRange & { isNullObject: boolean };
  setPrintTitleRows(address: string): void;
  /** Official: returns Range (not RangeAreas). */
  getPrintTitleColumnsOrNullObject(): ExcelRange & { isNullObject: boolean };
  setPrintTitleColumns(address: string): void;
}

export interface ExcelWorksheet {
  name: string;
  position: number;
  visibility: string;
  tabColor: string;
  showGridlines: boolean;
  showHeadings: boolean;
  protection: ExcelWorksheetProtection;
  freezePanes: ExcelFreezePanes;
  pageLayout: ExcelPageLayout;
  names: ExcelNamedItemCollection;
  load(props: string): void;
  getRange(address: string): ExcelRange;
  getUsedRangeOrNullObject(valuesOnly?: boolean): ExcelRange & { isNullObject: boolean };
  delete(): void;
  copy(positionType?: string, relativeTo?: ExcelWorksheet): ExcelWorksheet;
  tables: {
    items: ExcelTable[];
    add(address: string, hasHeaders: boolean): ExcelTable;
    getItem(name: string): ExcelTable;
    load(props: string): void;
  };
  charts: {
    items: ExcelChart[];
    add(type: string, sourceData: ExcelRange, seriesBy?: string): ExcelChart;
    getItem(name: string): ExcelChart;
    load(props: string): void;
  };
  shapes: import("./officeJsShapeFacade").ExcelShapeCollection;
}

export type {
  ExcelShape,
  ExcelShapeCollection,
  ExcelTextFrame,
} from "./officeJsShapeFacade";

export interface ExcelRequestContext {
  workbook: {
    name: string;
    load(props: string): void;
    names: ExcelNamedItemCollection;
    worksheets: {
      getActiveWorksheet(): ExcelWorksheet;
      getItem(name: string): ExcelWorksheet;
      add(name?: string): ExcelWorksheet;
      items: ExcelWorksheet[];
      load(props: string): void;
    };
    getSelectedRange(): ExcelRange;
  };
  sync(): Promise<void>;
}

declare global {
  interface Window {
    Office?: {
      onReady: (callback: (info: { host: string }) => void) => void;
    };
    Excel?: {
      run: ExcelRunFn;
      ChartType?: {
        columnClustered: string;
        line: string;
        barClustered: string;
        area: string;
        pie: string;
        xyscatter: string;
        doughnut: string;
        bubble: string;
        radar: string;
        lineMarkers: string;
      };
    };
  }
}

export function getExcelRun(): ExcelRunFn | null {
  if (typeof window === "undefined") return null;
  return window.Excel?.run ?? null;
}

export function normalizeMatrix(values: unknown): CellValue[][] {
  if (!Array.isArray(values)) return [];
  return values.map((row) => {
    if (!Array.isArray(row)) return [row as CellValue];
    return row.map((cell) => (cell === undefined ? null : (cell as CellValue)));
  });
}

export function normalizeFormulas(values: unknown): string[][] {
  if (!Array.isArray(values)) return [];
  return values.map((row) => {
    if (!Array.isArray(row)) return [String(row ?? "")];
    return row.map((cell) => String(cell ?? ""));
  });
}

export function firstNumberFormat(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (!Array.isArray(value) || value.length === 0) return null;
  const first = value[0];
  if (Array.isArray(first)) return first[0] != null ? String(first[0]) : null;
  return first != null ? String(first) : null;
}

export async function withExcel<T>(
  capability: string,
  fn: (context: ExcelRequestContext) => Promise<T>,
): Promise<HostResult<T>> {
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      capability,
      "office-js",
      "Excel.run is not available in this runtime",
      "Requires Microsoft Office Excel with Office.js loaded",
    );
  }
  try {
    return ok(await run(fn));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return unsupported(capability, "office-js", message, "Excel.run rejected the batch");
  }
}

export function loadRangeFormat(range: ExcelRange): void {
  range.load("address,numberFormat");
  range.format.load("horizontalAlignment,verticalAlignment,wrapText");
  range.format.font.load("name,size,bold,color");
  range.format.fill.load("color");
}

export function readFormatFromRange(range: ExcelRange): RangeFormat {
  return {
    fontName: range.format.font.name ?? null,
    fontSize: range.format.font.size ?? null,
    fontBold: range.format.font.bold ?? null,
    fontColor: range.format.font.color ?? null,
    fillColor: range.format.fill.color ?? null,
    numberFormat: firstNumberFormat(range.numberFormat),
    horizontalAlignment: range.format.horizontalAlignment ?? null,
    verticalAlignment: range.format.verticalAlignment ?? null,
    wrapText: range.format.wrapText ?? null,
  };
}
