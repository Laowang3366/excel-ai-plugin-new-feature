import type {
  ExcelPageLayout,
  ExcelPageLayoutZoomOptions,
  ExcelRequestContext,
  ExcelWorksheet,
} from "./officeJsRuntime";
import { getExcelRun } from "./officeJsRuntime";
import type {
  HostResult,
  PageOrientation,
  PagePaperSize,
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
} from "./types";
import { fail, ok, unsupported } from "./types";

const REQUIREMENT_EVIDENCE =
  "PageLayout.paperSize and PageLayout.zoom require ExcelApi 1.9";

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

function mapOrientation(value: string): PageOrientation {
  return String(value).toLowerCase().includes("landscape") ? "landscape" : "portrait";
}

function toOfficeOrientation(value: PageOrientation): string {
  return value === "landscape" ? "Landscape" : "Portrait";
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
  if (!("paperSize" in layout)) {
    throw new Error("PageLayout.paperSize is missing on host layout object");
  }
  if (!("zoom" in layout)) {
    throw new Error("PageLayout.zoom is missing on host layout object");
  }
  sheet.load("name");
  // Official: load whole zoom object, then read scale / fit fields after sync.
  layout.load(
    "orientation,centerHorizontally,centerVertically,printGridlines,printHeadings,blackAndWhite,topMargin,bottomMargin,leftMargin,rightMargin,paperSize,zoom",
  );
  const printArea = layout.getPrintAreaOrNullObject();
  const titleRows = layout.getPrintTitleRowsOrNullObject();
  const titleCols = layout.getPrintTitleColumnsOrNullObject();
  printArea.load("address");
  titleRows.load("address");
  titleCols.load("address");
  await context.sync();
  return {
    sheetName: requireLoadedString(sheet.name, "Worksheet.name"),
    orientation: mapOrientation(String(layout.orientation)),
    centerHorizontally: Boolean(layout.centerHorizontally),
    centerVertically: Boolean(layout.centerVertically),
    printGridlines: Boolean(layout.printGridlines),
    printHeadings: Boolean(layout.printHeadings),
    blackAndWhite: Boolean(layout.blackAndWhite),
    margins: {
      top: Number(layout.topMargin),
      bottom: Number(layout.bottomMargin),
      left: Number(layout.leftMargin),
      right: Number(layout.rightMargin),
    },
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
    if (input.margins) {
      if (input.margins.top != null) layout.topMargin = input.margins.top;
      if (input.margins.bottom != null) layout.bottomMargin = input.margins.bottom;
      if (input.margins.left != null) layout.leftMargin = input.margins.left;
      if (input.margins.right != null) layout.rightMargin = input.margins.right;
    }
    if (input.paperSize != null) {
      if (!("paperSize" in layout)) {
        throw new Error("PageLayout.paperSize is missing on host layout object");
      }
      layout.paperSize = PAPER_SIZE_TO_HOST[input.paperSize];
    }
    // Official: assign whole zoom options object (not zoom.scale / zoom.fit sub-property writes).
    const hasFit =
      input.fitToPagesWide !== undefined || input.fitToPagesTall !== undefined;
    if (input.zoomScale != null) {
      if (!("zoom" in layout)) {
        throw new Error("PageLayout.zoom is missing on host layout object");
      }
      layout.zoom = { scale: input.zoomScale };
    } else if (hasFit) {
      if (!("zoom" in layout)) {
        throw new Error("PageLayout.zoom is missing on host layout object");
      }
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
