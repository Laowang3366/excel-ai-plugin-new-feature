/** Office.js PageOrientation subset. */
export type PageOrientation = "portrait" | "landscape";

/** Public paperSize labels (exact lowercase input). */
export type PagePaperSize = "a3" | "a4" | "a5" | "letter" | "legal";

/** Public print order (maps to PageLayout.printOrder). */
export type PageOrder = "downThenOver" | "overThenDown";

export interface PageLayoutMargins {
  top: number;
  bottom: number;
  left: number;
  right: number;
  /** PageLayout.headerMargin (points). */
  header: number;
  /** PageLayout.footerMargin (points). */
  footer: number;
}

/** Default-page header/footer text sides (Office.js HeaderFooter). */
export interface PageLayoutTextSides {
  left: string;
  center: string;
  right: string;
}

/** Official Worksheet.pageLayout confirmed subset. */
export interface SheetPageLayoutInfo {
  sheetName: string;
  orientation: PageOrientation;
  centerHorizontally: boolean;
  centerVertically: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  blackAndWhite: boolean;
  /** Maps PageLayout.draftMode. */
  draft: boolean;
  /** Maps PageLayout.printOrder DownThenOver|OverThenDown. */
  pageOrder: PageOrder;
  /**
   * Maps PageLayout.firstPageNumber.
   * Host "" or null → null; finite integer >= 1 → number; other host values fail.
   */
  firstPageNumber: number | null;
  margins: PageLayoutMargins;
  /** defaultForAllPages left/center/rightHeader; "" when empty. */
  headers: PageLayoutTextSides;
  /** defaultForAllPages left/center/rightFooter; "" when empty. */
  footers: PageLayoutTextSides;
  /**
   * Percent scale from pageLayout.zoom.scale after load+sync.
   * Missing/non-finite scale → null (never coerce null → 0).
   * Fit-to-pages mode → null while fit fields may be set.
   */
  zoomScale: number | null;
  /** Host paperSize after load+sync; known values mapped to lowercase. */
  paperSize: string;
  /** From zoom.horizontalFitToPages; null when not in fit-wide mode. */
  fitToPagesWide: number | null;
  /** From zoom.verticalFitToPages; null when not in fit-tall mode. */
  fitToPagesTall: number | null;
  /** Unset print area/titles → null (via *OrNullObject). */
  printArea: string | null;
  printTitleRows: string | null;
  printTitleColumns: string | null;
  /** Manual horizontal page breaks as bare single-cell A1 (host order). */
  horizontalPageBreaks: string[];
  /** Manual vertical page breaks as bare single-cell A1 (host order). */
  verticalPageBreaks: string[];
}

export interface SheetPageLayoutUpdateInput {
  sheetName: string;
  orientation?: PageOrientation;
  centerHorizontally?: boolean;
  centerVertically?: boolean;
  printGridlines?: boolean;
  printHeadings?: boolean;
  blackAndWhite?: boolean;
  /** Must accept false (draftMode). */
  draft?: boolean;
  pageOrder?: PageOrder;
  /** Finite integer >= 1 only; clear/auto unsupported. */
  firstPageNumber?: number;
  margins?: Partial<PageLayoutMargins>;
  /** Partial default-page header sides; "" clears that side. */
  headers?: Partial<PageLayoutTextSides>;
  /** Partial default-page footer sides; "" clears that side. */
  footers?: Partial<PageLayoutTextSides>;
  zoomScale?: number;
  /** Exact lowercase; maps to host A3|A4|A5|Letter|Legal. */
  paperSize?: PagePaperSize;
  /** Finite integer 1..32767; mutually exclusive with zoomScale. */
  fitToPagesWide?: number;
  /** Finite integer 1..32767; mutually exclusive with zoomScale. */
  fitToPagesTall?: number;
  /** Non-empty string only; clear is not a proven Office.js no-arg contract. */
  printArea?: string;
  printTitleRows?: string;
  printTitleColumns?: string;
  /** true: clear both manual break collections before appends in this request. */
  clearPageBreaks?: boolean;
  /** Append manual horizontal breaks (bare A1); [] is no-op, not clear. */
  horizontalPageBreaks?: string[];
  /** Append manual vertical breaks (bare A1); [] is no-op, not clear. */
  verticalPageBreaks?: string[];
}
