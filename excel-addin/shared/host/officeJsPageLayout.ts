import type {
  ExcelPageLayout,
  ExcelPageLayoutZoomOptions,
  ExcelRequestContext,
  ExcelWorksheet,
} from "./officeJsRuntime";
import {
  applyDefaultHeadersFooters,
  loadDefaultHeadersFooters,
  readDefaultHeadersFooters,
} from "./officeJsPageLayoutHeadersFooters";
import { getExcelRun } from "./officeJsRuntime";
import type {
  HostResult,
  PageOrder,
  PageOrientation,
  PagePaperSize,
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
} from "./types";
import { fail, ok, unsupported } from "./types";

const REQUIREMENT_EVIDENCE =
  "PageLayout paperSize/zoom/draftMode/printOrder/firstPageNumber/headerMargin/footerMargin/headersFooters.defaultForAllPages require ExcelApi 1.9";

const PAPER_SIZE_TO_HOST: Record<PagePaperSize, string> = {
  a3: "A3",
  a4: "A4",
  a5: "A5",
  letter: "Letter",
  legal: "Legal",
};

const HOST_PAPER_SIZE_TO_PUBLIC: Record<string, PagePaperSize> = {
  A3: "a3",
  A4: "a4",
  A5: "a5",
  Letter: "letter",
  Legal: "legal",
};

const PAGE_ORDER_TO_HOST: Record<PageOrder, string> = {
  downThenOver: "DownThenOver",
  overThenDown: "OverThenDown",
};

function mapOrientation(value: string): PageOrientation {
  return String(value).toLowerCase().includes("landscape") ? "landscape" : "portrait";
}

function toOfficeOrientation(value: PageOrientation): string {
  return value === "landscape" ? "Landscape" : "Portrait";
}

function mapPageOrderFromHost(value: unknown): PageOrder {
  const raw = String(value ?? "");
  if (raw === "OverThenDown" || raw.toLowerCase() === "overthendown") return "overThenDown";
  if (raw === "DownThenOver" || raw.toLowerCase() === "downthenover") return "downThenOver";
  throw new Error(`PageLayout.printOrder has unknown host value: ${raw}`);
}

function addressOrNull(areas: { isNullObject: boolean; address: string }): string | null {
  return areas.isNullObject ? null : areas.address;
}

/** Read zoom.scale only; null/undefined/non-finite → null (not 0). */
function readZoomScale(zoom: ExcelPageLayoutZoomOptions | undefined): number | null {
  const scale = zoom?.scale;
  if (scale == null || typeof scale !== "number" || !Number.isFinite(scale)) return null;
  return scale;
}

/** Host fit page count; non-finite / missing → null. */
function readFitPages(value: unknown): number | null {
  if (value == null || typeof value !== "number" || !Number.isFinite(value)) return null;
  return value;
}

/**
 * Host firstPageNumber: "" or null → null; finite integer >= 1 → number;
 * other values are ordinary failures.
 */
function readFirstPageNumber(value: unknown): number | null {
  if (value === null || value === "") return null;
  if (typeof value === "number" && Number.isFinite(value) && Number.isInteger(value) && value >= 1) {
    return value;
  }
  throw new Error("PageLayout.firstPageNumber is not a finite integer >= 1 or empty/null");
}

function mapPaperSizeFromHost(value: unknown): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error("paperSize is not a loaded non-empty string");
  }
  return HOST_PAPER_SIZE_TO_PUBLIC[value] ?? value;
}

function requireLoadedString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded non-empty string`);
  }
  return value;
}

function requireMember(layout: ExcelPageLayout, key: string): void {
  if (!(key in layout)) {
    throw new Error(`PageLayout.${key} is missing on host layout object`);
  }
}

/** Official precheck before any Excel.run / pageLayout access. */
export function isExcelApi19SupportedForPageLayout(): boolean {
  const office = (
    globalThis as unknown as {
      Office?: {
        context?: {
          requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
        };
      };
    }
  ).Office;
  const isSetSupported = office?.context?.requirements?.isSetSupported;
  if (typeof isSetSupported !== "function") return false;
  try {
    return isSetSupported.call(office!.context!.requirements, "ExcelApi", "1.9");
  } catch {
    return false;
  }
}

function pageLayoutUnsupported(capability: string, reason: string): HostResult<SheetPageLayoutInfo> {
  return unsupported(capability, "office-js", reason, REQUIREMENT_EVIDENCE);
}

const EXCEL_RUN_EVIDENCE = "Requires Office.js Excel.run";

/**
 * Local pageLayout runner: after ExcelApi 1.9 precheck, missing Excel.run is typed unsupported;
 * host/runtime errors inside the batch are ordinary FailResult with capability + host.
 * Does not use withExcel (which always marks catch as unsupported).
 */
async function runPageLayout<T>(
  capability: string,
  fn: (context: ExcelRequestContext) => Promise<T>,
): Promise<HostResult<T>> {
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      capability,
      "office-js",
      "Excel.run is not available in this runtime",
      EXCEL_RUN_EVIDENCE,
    );
  }
  try {
    return ok(await run(fn));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(capability, "office-js", message);
  }
}

async function readLayout(
  sheet: ExcelWorksheet,
  layout: ExcelPageLayout,
  context: ExcelRequestContext,
): Promise<SheetPageLayoutInfo> {
  // Post-precheck: missing members are ordinary failures, not requirement-set unsupported.
  requireMember(layout, "paperSize");
  requireMember(layout, "zoom");
  requireMember(layout, "draftMode");
  requireMember(layout, "printOrder");
  requireMember(layout, "firstPageNumber");
  requireMember(layout, "headerMargin");
  requireMember(layout, "footerMargin");
  requireMember(layout, "headersFooters");
  if (!layout.headersFooters || typeof layout.headersFooters !== "object") {
    throw new Error("PageLayout.headersFooters is missing on host layout object");
  }
  if (
    !("defaultForAllPages" in layout.headersFooters) ||
    layout.headersFooters.defaultForAllPages == null ||
    typeof layout.headersFooters.defaultForAllPages !== "object"
  ) {
    throw new Error(
      "PageLayout.headersFooters.defaultForAllPages is missing on host layout object",
    );
  }
  sheet.load("name");
  // Official: load whole zoom object, then read scale / fit fields after sync.
  layout.load(
    "orientation,centerHorizontally,centerVertically,printGridlines,printHeadings,blackAndWhite,draftMode,printOrder,firstPageNumber,topMargin,bottomMargin,leftMargin,rightMargin,headerMargin,footerMargin,paperSize,zoom",
  );
  loadDefaultHeadersFooters(layout.headersFooters.defaultForAllPages);
  const printArea = layout.getPrintAreaOrNullObject();
  const titleRows = layout.getPrintTitleRowsOrNullObject();
  const titleCols = layout.getPrintTitleColumnsOrNullObject();
  printArea.load("address");
  titleRows.load("address");
  titleCols.load("address");
  await context.sync();
  const headersFootersText = readDefaultHeadersFooters(layout.headersFooters.defaultForAllPages);
  return {
    sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
    orientation: mapOrientation(String(layout.orientation)),
    centerHorizontally: Boolean(layout.centerHorizontally),
    centerVertically: Boolean(layout.centerVertically),
    printGridlines: Boolean(layout.printGridlines),
    printHeadings: Boolean(layout.printHeadings),
    blackAndWhite: Boolean(layout.blackAndWhite),
    draft: Boolean(layout.draftMode),
    pageOrder: mapPageOrderFromHost(layout.printOrder),
    firstPageNumber: readFirstPageNumber(layout.firstPageNumber),
    margins: {
      top: Number(layout.topMargin),
      bottom: Number(layout.bottomMargin),
      left: Number(layout.leftMargin),
      right: Number(layout.rightMargin),
      header: Number(layout.headerMargin),
      footer: Number(layout.footerMargin),
    },
    headers: headersFootersText.headers,
    footers: headersFootersText.footers,
    zoomScale: readZoomScale(layout.zoom),
    paperSize: mapPaperSizeFromHost(layout.paperSize),
    fitToPagesWide: readFitPages(layout.zoom?.horizontalFitToPages),
    fitToPagesTall: readFitPages(layout.zoom?.verticalFitToPages),
    printArea: addressOrNull(printArea),
    printTitleRows: addressOrNull(titleRows),
    printTitleColumns: addressOrNull(titleCols),
  };
}

export async function officeJsGetSheetPageLayout(
  sheetName: string,
): Promise<HostResult<SheetPageLayoutInfo>> {
  if (!isExcelApi19SupportedForPageLayout()) {
    return pageLayoutUnsupported(
      "sheet.pageLayout.get",
      "ExcelApi 1.9 is not supported in this host (Office.context.requirements.isSetSupported)",
    );
  }
  return runPageLayout("sheet.pageLayout.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    return readLayout(sheet, sheet.pageLayout, context);
  });
}

export async function officeJsSetSheetPageLayout(
  input: SheetPageLayoutUpdateInput,
): Promise<HostResult<SheetPageLayoutInfo>> {
  if (!isExcelApi19SupportedForPageLayout()) {
    return pageLayoutUnsupported(
      "sheet.pageLayout.set",
      "ExcelApi 1.9 is not supported in this host (Office.context.requirements.isSetSupported)",
    );
  }
  return runPageLayout("sheet.pageLayout.set", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    const layout = sheet.pageLayout;
    if (input.orientation != null) layout.orientation = toOfficeOrientation(input.orientation);
    if (input.centerHorizontally != null) layout.centerHorizontally = input.centerHorizontally;
    if (input.centerVertically != null) layout.centerVertically = input.centerVertically;
    if (input.printGridlines != null) layout.printGridlines = input.printGridlines;
    if (input.printHeadings != null) layout.printHeadings = input.printHeadings;
    if (input.blackAndWhite != null) layout.blackAndWhite = input.blackAndWhite;
    if (input.draft !== undefined) {
      requireMember(layout, "draftMode");
      layout.draftMode = input.draft;
    }
    if (input.pageOrder != null) {
      requireMember(layout, "printOrder");
      layout.printOrder = PAGE_ORDER_TO_HOST[input.pageOrder];
    }
    if (input.firstPageNumber !== undefined) {
      requireMember(layout, "firstPageNumber");
      layout.firstPageNumber = input.firstPageNumber;
    }
    if (input.margins) {
      if (input.margins.top != null) layout.topMargin = input.margins.top;
      if (input.margins.bottom != null) layout.bottomMargin = input.margins.bottom;
      if (input.margins.left != null) layout.leftMargin = input.margins.left;
      if (input.margins.right != null) layout.rightMargin = input.margins.right;
      if (input.margins.header != null) {
        requireMember(layout, "headerMargin");
        layout.headerMargin = input.margins.header;
      }
      if (input.margins.footer != null) {
        requireMember(layout, "footerMargin");
        layout.footerMargin = input.margins.footer;
      }
    }
    if (input.headers !== undefined || input.footers !== undefined) {
      requireMember(layout, "headersFooters");
      if (!layout.headersFooters || typeof layout.headersFooters !== "object") {
        throw new Error("PageLayout.headersFooters is missing on host layout object");
      }
      if (
        !("defaultForAllPages" in layout.headersFooters) ||
        layout.headersFooters.defaultForAllPages == null ||
        typeof layout.headersFooters.defaultForAllPages !== "object"
      ) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages is missing on host layout object",
        );
      }
      applyDefaultHeadersFooters(
        layout.headersFooters.defaultForAllPages,
        input.headers,
        input.footers,
      );
    }
    if (input.paperSize != null) {
      requireMember(layout, "paperSize");
      layout.paperSize = PAPER_SIZE_TO_HOST[input.paperSize];
    }
    // Official: assign whole zoom options object (not zoom.scale / zoom.fit sub-property writes).
    const hasFit =
      input.fitToPagesWide !== undefined || input.fitToPagesTall !== undefined;
    if (input.zoomScale != null) {
      requireMember(layout, "zoom");
      layout.zoom = { scale: input.zoomScale };
    } else if (hasFit) {
      requireMember(layout, "zoom");
      const zoom: ExcelPageLayoutZoomOptions = {};
      if (input.fitToPagesWide !== undefined) zoom.horizontalFitToPages = input.fitToPagesWide;
      if (input.fitToPagesTall !== undefined) zoom.verticalFitToPages = input.fitToPagesTall;
      layout.zoom = zoom;
    }
    // setPrintArea/title require non-empty string; clear is not a proven no-arg contract.
    if (input.printArea != null) layout.setPrintArea(input.printArea);
    if (input.printTitleRows != null) layout.setPrintTitleRows(input.printTitleRows);
    if (input.printTitleColumns != null) layout.setPrintTitleColumns(input.printTitleColumns);
    return readLayout(sheet, layout, context);
  });
}
