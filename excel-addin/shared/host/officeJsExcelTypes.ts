import type { CellValue } from "./types";
import type { ExcelPageBreakCollection } from "./officeJsPageBreakTypes";

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

/** ExcelApi 1.2 Range.format.protection */
export interface ExcelRangeFormatProtection {
  locked: boolean;
  load(props: string): void;
}

export interface ExcelRangeFormat {
  protection: ExcelRangeFormatProtection;
  font: ExcelFont;
  fill: ExcelFill;
  columnWidth: number | null;
  rowHeight: number | null;
  horizontalAlignment: string;
  verticalAlignment: string;
  wrapText: boolean;
  /** ExcelApi 1.2. */
  autofitColumns(): void;
  /** ExcelApi 1.2. */
  autofitRows(): void;
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
  decimal?: {
    formula1?: string | number;
    formula2?: string | number;
    operator?: string;
  };
  date?: {
    formula1?: string | number;
    formula2?: string | number;
    operator?: string;
  };
  time?: {
    formula1?: string | number;
    formula2?: string | number;
    operator?: string;
  };
  textLength?: {
    formula1?: string | number;
    formula2?: string | number;
    operator?: string;
  };
  custom?: { formula?: string };
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
  /** ExcelApi 1.7: host Base64 PNG; no width/height params. */
  getImage(): { value: string };
  /** ExcelApi 1.1. */
  insert(shift: "Down" | "Right"): ExcelRange;
  /** ExcelApi 1.1. */
  delete(shift: "Up" | "Left"): void;
  getSpillingToRange(): ExcelRange;
  /** ExcelApi 1.12: null object when cell is not a spill parent. */
  getSpillingToRangeOrNullObject(): ExcelRange;
  getSurroundingRegion(): ExcelRange;
  getCurrentArray(): ExcelRange;
  getCell(row: number, column: number): ExcelRange;
}

export interface ExcelTableAutoFilter {
  enabled: boolean;
  apply(range: ExcelRange | unknown, columnIndex: number, criteria: unknown): void;
  clearCriteria(): void;
  load(props: string): void;
}

export interface ExcelTableSortField {
  key?: number;
  ascending?: boolean;
  load?(props: string): void;
}

export interface ExcelTableSort {
  fields: {
    load(props: string): void;
    items?: ExcelTableSortField[];
  };
  apply(fields: unknown, matchCase?: boolean): void;
  clear(): void;
  load?(props: string): void;
}

export interface ExcelTable {
  name: string;
  /** ExcelApi 1.3. */
  showBandedColumns: boolean;
  /** ExcelApi 1.3. */
  showBandedRows: boolean;
  /** ExcelApi 1.3. */
  showFirstColumn: boolean;
  /** ExcelApi 1.3. */
  showLastColumn: boolean;
  showHeaders: boolean;
  showFilterButton: boolean;
  showTotals: boolean;
  style: string;
  /** ExcelApi 1.2+. */
  autoFilter: ExcelTableAutoFilter;
  /** ExcelApi 1.2+. */
  sort: ExcelTableSort;
  getRange(): ExcelRange;
  /** ExcelApi 1.13; new range must overlap and keep the header row. */
  resize(newRange: ExcelRange | string): void;
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
  /** Official: Chart.setData(sourceData: Range, seriesBy?: ChartSeriesBy). */
  setData(sourceData: ExcelRange, seriesBy?: string): void;
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
  /** ExcelApi 1.9 manual page breaks only. */
  horizontalPageBreaks: ExcelPageBreakCollection;
  verticalPageBreaks: ExcelPageBreakCollection;
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
      context?: {
        requirements?: {
          isSetSupported(name: string, version: string): boolean;
        };
      };
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

