/**
 * workbook.template.capture — shallow desktop capture/inspect parity.
 * Single Excel.run; fixed small sync batches (not O(sheetCount)).
 */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelRange, ExcelWorksheet } from "./officeJsExcelTypes";
import { getExcelRun } from "./officeJsRuntime";
import {
  parseTemplatePrintSnapshot,
  queueTemplatePrintLoads,
  type TemplatePrintLoadBundle,
} from "./officeJsTemplatePrint";
import { requireExcelApi19ForTemplateCapture } from "./officeJsTemplateRequirements";
import {
  nullableBoolean,
  nullableHexColor,
  nullableNonEmptyString,
  nullablePositiveFinite,
  requireBoolean,
  requireNonEmptyString,
  requirePositiveInt,
  splitSheetQualifiedAddress,
} from "./officeJsTemplateReadback";
import type {
  WorkbookTemplateBaseStyle,
  WorkbookTemplateCaptureInfo,
  WorkbookTemplateCapturedSheet,
  WorkbookTemplateHeaderStyle,
} from "./workbookTemplateTypes";
import { WORKBOOK_TEMPLATE_MAX_SHEETS } from "./workbookTemplateTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const CAPABILITY = "workbook.template.capture";

function assertCaptureSheetSurface(sheet: ExcelWorksheet): void {
  if (typeof sheet.getUsedRangeOrNullObject !== "function") {
    throw new Error("Worksheet.getUsedRangeOrNullObject is missing");
  }
  if (typeof sheet.getRange !== "function") throw new Error("Worksheet.getRange is missing");
  if (typeof sheet.load !== "function") throw new Error("Worksheet.load is missing");
  if (!("pageLayout" in sheet) || sheet.pageLayout == null) {
    throw new Error("Worksheet.pageLayout is missing");
  }
}

function headerRowAddress(usedAddress: string, columnCount: number): string {
  const bare = splitSheetQualifiedAddress(usedAddress).bare;
  const start = bare.includes(":") ? bare.split(":")[0]! : bare;
  const cell = parseA1Cell(start);
  if (!cell) throw new Error(`Cannot parse usedRange address for header: ${usedAddress}`);
  const endCol = cell.col + Math.max(1, columnCount) - 1;
  return `${toA1(cell.row, cell.col)}:${toA1(cell.row, endCol)}`;
}

function parseBaseStyle(
  used: ExcelRange,
  limitations: string[],
): WorkbookTemplateBaseStyle {
  if (!used.format?.font) throw new Error("UsedRange.format.font is missing");
  const fontName = nullableNonEmptyString(used.format.font.name, "base.font.name");
  const fontSize = nullablePositiveFinite(used.format.font.size, "base.font.size");
  const fontColor = nullableHexColor(used.format.font.color, "base.font.color");
  if (fontName === null) limitations.push("baseStyle.fontName mixed/unavailable (null)");
  if (fontSize === null) limitations.push("baseStyle.fontSize mixed/unavailable (null)");
  if (fontColor === null) limitations.push("baseStyle.fontColor mixed/unavailable (null)");
  return { fontName, fontSize, fontColor };
}

function parseHeaderStyle(
  header: ExcelRange,
  limitations: string[],
): WorkbookTemplateHeaderStyle {
  if (!header.format?.font || !header.format.fill) {
    throw new Error("HeaderRange.format font/fill is missing");
  }
  const fillColor = nullableHexColor(header.format.fill.color, "header.fill.color");
  const fontColor = nullableHexColor(header.format.font.color, "header.font.color");
  const bold = nullableBoolean(header.format.font.bold, "header.font.bold");
  const rowHeight = nullablePositiveFinite(header.format.rowHeight, "header.rowHeight");
  if (fillColor === null) limitations.push("headerStyle.fillColor mixed/unavailable (null)");
  if (fontColor === null) limitations.push("headerStyle.fontColor mixed/unavailable (null)");
  if (bold === null) limitations.push("headerStyle.bold mixed/unavailable (null)");
  if (rowHeight === null) limitations.push("headerStyle.rowHeight mixed/unavailable (null)");
  return { fillColor, fontColor, bold, rowHeight };
}

type QueuedSheet = {
  sheet: ExcelWorksheet;
  used: ExcelRange & { isNullObject?: unknown };
  print: TemplatePrintLoadBundle;
};

type StyleQueued = QueuedSheet & {
  empty: boolean;
  sheetName: string;
  usedRange: string | null;
  rows: number;
  columns: number;
  header?: ExcelRange;
  sheetLimitations: string[];
};

export async function officeJsCaptureWorkbookTemplate(): Promise<
  HostResult<WorkbookTemplateCaptureInfo>
> {
  const gate = requireExcelApi19ForTemplateCapture(CAPABILITY);
  if (gate) return gate;

  const run = getExcelRun();
  if (!run) {
    return unsupported(
      CAPABILITY,
      "office-js",
      "Excel.run is not available in this runtime",
      "Requires Office.js Excel.run",
    );
  }

  try {
    const info = await run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      context.workbook.load("name");
      // sync #1 — sheet list + workbook name
      await context.sync();

      const items = sheets.items;
      if (!Array.isArray(items)) throw new Error("WorksheetCollection.items is not an array");
      if (items.length > WORKBOOK_TEMPLATE_MAX_SHEETS) {
        throw new Error(
          `resource-limit: workbook exceeds ${WORKBOOK_TEMPLATE_MAX_SHEETS} worksheets`,
        );
      }

      const workbookName = requireNonEmptyString(context.workbook.name, "Workbook.name");
      const queued: QueuedSheet[] = [];
      for (const sheet of items) {
        assertCaptureSheetSurface(sheet);
        sheet.load("name");
        const used = sheet.getUsedRangeOrNullObject(false) as ExcelRange & {
          isNullObject?: unknown;
        };
        if (typeof used.load !== "function") {
          throw new Error("UsedRange.load is missing");
        }
        used.load("isNullObject,address,rowCount,columnCount");
        const print = queueTemplatePrintLoads(sheet);
        queued.push({ sheet, used, print });
      }
      // sync #2 — used dims + print snapshot props for all sheets
      await context.sync();

      const styleQueued: StyleQueued[] = [];
      for (const q of queued) {
        const sheetName = requireNonEmptyString(q.sheet.name, "Worksheet.name");
        const sheetLimitations: string[] = [];
        const isNull = requireBoolean(q.used.isNullObject, "UsedRange.isNullObject");
        if (isNull) {
          styleQueued.push({
            ...q,
            empty: true,
            sheetName,
            usedRange: null,
            rows: 0,
            columns: 0,
            sheetLimitations: ["empty sheet: usedRange/baseStyle/headerStyle null"],
          });
          continue;
        }
        const addr = requireNonEmptyString(q.used.address, "UsedRange.address");
        const usedRange = splitSheetQualifiedAddress(addr).bare.replace(/\$/g, "").toUpperCase();
        const rows = requirePositiveInt(q.used.rowCount, "UsedRange.rowCount");
        const columns = requirePositiveInt(q.used.columnCount, "UsedRange.columnCount");
        if (!q.used.format?.font || typeof q.used.format.font.load !== "function") {
          throw new Error("UsedRange.format.font.load is missing");
        }
        q.used.format.font.load("name,size,color");
        const header = q.sheet.getRange(headerRowAddress(addr, columns));
        if (!header.format?.font || typeof header.format.font.load !== "function") {
          throw new Error("HeaderRange.format.font.load is missing");
        }
        if (!header.format.fill || typeof header.format.fill.load !== "function") {
          throw new Error("HeaderRange.format.fill.load is missing");
        }
        if (typeof header.format.load !== "function") {
          throw new Error("HeaderRange.format.load is missing");
        }
        header.format.font.load("bold,color");
        header.format.fill.load("color");
        header.format.load("rowHeight");
        styleQueued.push({
          ...q,
          empty: false,
          sheetName,
          usedRange,
          rows,
          columns,
          header,
          sheetLimitations,
        });
      }
      // sync #3 — base/header styles for non-empty sheets only
      if (styleQueued.some((s) => !s.empty)) {
        await context.sync();
      }

      const captured: WorkbookTemplateCapturedSheet[] = [];
      for (const s of styleQueued) {
        let baseStyle: WorkbookTemplateBaseStyle | null = null;
        let headerStyle: WorkbookTemplateHeaderStyle | null = null;
        if (!s.empty && s.header) {
          baseStyle = parseBaseStyle(s.used, s.sheetLimitations);
          headerStyle = parseHeaderStyle(s.header, s.sheetLimitations);
        }
        const print = parseTemplatePrintSnapshot(s.print);
        captured.push({
          name: s.sheetName,
          usedRange: s.usedRange,
          rows: s.rows,
          columns: s.columns,
          baseStyle,
          headerStyle,
          print,
          limitations: s.sheetLimitations,
        });
      }

      return {
        template: {
          version: 1 as const,
          capturedFrom: workbookName,
          capturedAt: new Date().toISOString(),
          sheets: captured,
        },
        sheetCount: captured.length,
        limitations: [
          "Shallow snapshot only — not a full theme/table-style/CF/DV dump; not a replayable template package",
          "Print snapshot is a template subset (area/orientation/paperSize/fit/titles/center header/footer); not full pageLayout",
          "Not real Excel sideload verified",
        ],
      } satisfies WorkbookTemplateCaptureInfo;
    });
    return ok(info);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.startsWith("resource-limit:")) {
      return fail(CAPABILITY, "office-js", message, "max 500 worksheets");
    }
    return fail(CAPABILITY, "office-js", message);
  }
}
