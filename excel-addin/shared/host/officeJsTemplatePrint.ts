/**
 * Capture-only template print snapshot helpers (strict; no full pageLayout / page breaks).
 * Queue loads, shared sync, then parse — no nested Excel.run.
 */
import type { ExcelPageLayout, ExcelRange, ExcelWorksheet } from "./officeJsExcelTypes";
import {
  normalizeRangeAddressForCompare,
  nullableNonNegativeInt,
  nullableOrientation,
  nullablePaperSizeToken,
  nullableStringAllowEmpty,
  requireBoolean,
  requireNonEmptyString,
} from "./officeJsTemplateReadback";
import type { WorkbookTemplatePrintSnapshot } from "./workbookTemplateTypes";

type NullRange = ExcelRange & { isNullObject?: unknown; address?: unknown };

export type TemplatePrintLoadBundle = {
  layout: ExcelPageLayout;
  printArea: NullRange;
  titleRows: NullRange;
  titleCols: NullRange;
  headerCenter: { value: () => unknown };
  footerCenter: { value: () => unknown };
};

function requirePageLayout(sheet: ExcelWorksheet): ExcelPageLayout {
  if (!("pageLayout" in sheet) || sheet.pageLayout == null) {
    throw new Error("Worksheet.pageLayout is missing");
  }
  const layout = sheet.pageLayout as ExcelPageLayout;
  if (typeof layout.load !== "function") {
    throw new Error("PageLayout.load is missing");
  }
  if (typeof layout.getPrintAreaOrNullObject !== "function") {
    throw new Error("PageLayout.getPrintAreaOrNullObject is missing");
  }
  if (typeof layout.getPrintTitleRowsOrNullObject !== "function") {
    throw new Error("PageLayout.getPrintTitleRowsOrNullObject is missing");
  }
  if (typeof layout.getPrintTitleColumnsOrNullObject !== "function") {
    throw new Error("PageLayout.getPrintTitleColumnsOrNullObject is missing");
  }
  return layout;
}

function requireDefaultHeaders(layout: ExcelPageLayout): {
  centerHeader: unknown;
  centerFooter: unknown;
} {
  if (!("headersFooters" in layout) || layout.headersFooters == null) {
    throw new Error("PageLayout.headersFooters is missing");
  }
  const hf = layout.headersFooters as {
    defaultForAllPages?: {
      centerHeader?: unknown;
      centerFooter?: unknown;
      load?: (props: string) => void;
    };
  };
  if (!hf.defaultForAllPages || typeof hf.defaultForAllPages !== "object") {
    throw new Error("PageLayout.headersFooters.defaultForAllPages is missing");
  }
  const def = hf.defaultForAllPages;
  if (typeof def.load === "function") {
    def.load("centerHeader,centerFooter");
  }
  return {
    centerHeader: def,
    centerFooter: def,
  };
}

/** Queue print-related loads for a sheet (no sync). */
export function queueTemplatePrintLoads(sheet: ExcelWorksheet): TemplatePrintLoadBundle {
  const layout = requirePageLayout(sheet);
  layout.load("orientation,paperSize,zoom");
  const headers = requireDefaultHeaders(layout);
  const printArea = layout.getPrintAreaOrNullObject() as NullRange;
  const titleRows = layout.getPrintTitleRowsOrNullObject() as NullRange;
  const titleCols = layout.getPrintTitleColumnsOrNullObject() as NullRange;
  if (typeof printArea.load !== "function") {
    throw new Error("printArea.load is missing");
  }
  if (typeof titleRows.load !== "function") {
    throw new Error("titleRows.load is missing");
  }
  if (typeof titleCols.load !== "function") {
    throw new Error("titleCols.load is missing");
  }
  printArea.load("isNullObject,address");
  titleRows.load("isNullObject,address");
  titleCols.load("isNullObject,address");
  return {
    layout,
    printArea,
    titleRows,
    titleCols,
    headerCenter: {
      value: () =>
        (headers.centerHeader as { centerHeader?: unknown }).centerHeader,
    },
    footerCenter: {
      value: () =>
        (headers.centerFooter as { centerFooter?: unknown }).centerFooter,
    },
  };
}

function parseNullAddress(range: NullRange, field: string): string | null {
  const isNull = requireBoolean(range.isNullObject, `${field}.isNullObject`);
  if (isNull) return null;
  const addr = requireNonEmptyString(range.address, `${field}.address`);
  return normalizeRangeAddressForCompare(addr);
}

function readFit(zoom: unknown, key: "horizontalFitToPages" | "verticalFitToPages", field: string): number | null {
  if (zoom === null) return null;
  if (zoom == null || typeof zoom !== "object") {
    throw new Error(`${field} zoom is not a loaded object or null`);
  }
  const raw = (zoom as Record<string, unknown>)[key];
  return nullableNonNegativeInt(raw === undefined ? null : raw, field);
}

/** Parse after shared sync — strict scalars, no String/Boolean/Number coercion. */
export function parseTemplatePrintSnapshot(
  bundle: TemplatePrintLoadBundle,
): WorkbookTemplatePrintSnapshot {
  const layout = bundle.layout as ExcelPageLayout & {
    orientation?: unknown;
    paperSize?: unknown;
    zoom?: unknown;
  };
  const orientation = nullableOrientation(layout.orientation ?? null, "print.orientation");
  const paperSize = nullablePaperSizeToken(layout.paperSize ?? null, "print.paperSize");
  const zoom = layout.zoom;
  const fitToPagesWide = readFit(zoom, "horizontalFitToPages", "print.fitToPagesWide");
  const fitToPagesTall = readFit(zoom, "verticalFitToPages", "print.fitToPagesTall");
  return {
    area: parseNullAddress(bundle.printArea, "print.area"),
    orientation,
    paperSize,
    fitToPagesWide,
    fitToPagesTall,
    repeatRows: parseNullAddress(bundle.titleRows, "print.repeatRows"),
    repeatColumns: parseNullAddress(bundle.titleCols, "print.repeatColumns"),
    header: nullableStringAllowEmpty(bundle.headerCenter.value(), "print.header"),
    footer: nullableStringAllowEmpty(bundle.footerCenter.value(), "print.footer"),
  };
}
