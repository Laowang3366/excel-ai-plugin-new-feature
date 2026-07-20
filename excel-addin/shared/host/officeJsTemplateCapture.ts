/**
 * workbook.template.capture — shallow desktop capture/inspect parity.
 * Single Excel.run batch; uses context-bound pageLayout reader (no nested Excel.run).
 */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";
import { readSheetPageLayoutInContext } from "./officeJsPageLayout";
import { getExcelRun } from "./officeJsRuntime";
import { requireExcelApi19ForTemplateCapture } from "./officeJsTemplateRequirements";
import {
  nullableBoolean,
  nullableFiniteNumber,
  nullableHexColor,
  nullableString,
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
  WorkbookTemplatePrintSnapshot,
} from "./workbookTemplateTypes";
import { WORKBOOK_TEMPLATE_MAX_SHEETS } from "./workbookTemplateTypes";
import type { HostResult, SheetPageLayoutInfo } from "./types";
import { fail, ok, unsupported } from "./types";

const CAPABILITY = "workbook.template.capture";

function printFromPageLayout(layout: SheetPageLayoutInfo): WorkbookTemplatePrintSnapshot {
  return {
    area: layout.printArea,
    orientation: layout.orientation,
    paperSize: layout.paperSize,
    fitToPagesWide: layout.fitToPagesWide,
    fitToPagesTall: layout.fitToPagesTall,
    repeatRows: layout.printTitleRows,
    repeatColumns: layout.printTitleColumns,
    header: layout.headers.center,
    footer: layout.footers.center,
  };
}

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

async function readBaseStyle(
  used: ExcelRange,
  context: ExcelRequestContext,
  limitations: string[],
): Promise<WorkbookTemplateBaseStyle> {
  if (!used.format?.font || typeof used.format.font.load !== "function") {
    throw new Error("UsedRange.format.font.load is missing");
  }
  used.format.font.load("name,size,color");
  await context.sync();
  const fontName = nullableString(used.format.font.name, "base.font.name");
  const fontSize = nullableFiniteNumber(used.format.font.size, "base.font.size");
  const fontColor = nullableHexColor(used.format.font.color, "base.font.color");
  if (fontName === null) limitations.push("baseStyle.fontName mixed/unavailable (null)");
  if (fontSize === null) limitations.push("baseStyle.fontSize mixed/unavailable (null)");
  if (fontColor === null) limitations.push("baseStyle.fontColor mixed/unavailable (null)");
  return { fontName, fontSize, fontColor };
}

async function readHeaderStyle(
  sheet: ExcelWorksheet,
  used: ExcelRange,
  context: ExcelRequestContext,
  limitations: string[],
): Promise<WorkbookTemplateHeaderStyle> {
  const addr = requireNonEmptyString(used.address, "UsedRange.address");
  const bare = splitSheetQualifiedAddress(addr).bare;
  const start = bare.includes(":") ? bare.split(":")[0]! : bare;
  used.load("columnCount");
  await context.sync();
  const cols = requirePositiveInt(used.columnCount, "UsedRange.columnCount");
  const cell = parseA1Cell(start);
  if (!cell) throw new Error(`Cannot parse usedRange for header: ${addr}`);
  const headerAddr = `${toA1(cell.row, cell.col)}:${toA1(cell.row, cell.col + cols - 1)}`;
  const header = sheet.getRange(headerAddr);
  if (!header.format?.font || typeof header.format.font.load !== "function") {
    throw new Error("HeaderRange.format.font.load is missing");
  }
  if (!header.format.fill || typeof header.format.fill.load !== "function") {
    throw new Error("HeaderRange.format.fill.load is missing");
  }
  header.format.font.load("bold,color");
  header.format.fill.load("color");
  header.format.load("rowHeight");
  await context.sync();

  const fillColor = nullableHexColor(header.format.fill.color, "header.fill.color");
  const fontColor = nullableHexColor(header.format.font.color, "header.font.color");
  const bold = nullableBoolean(header.format.font.bold, "header.font.bold");
  const rowHeight = nullableFiniteNumber(header.format.rowHeight, "header.rowHeight");
  if (fillColor === null) limitations.push("headerStyle.fillColor mixed/unavailable (null)");
  if (fontColor === null) limitations.push("headerStyle.fontColor mixed/unavailable (null)");
  if (bold === null) limitations.push("headerStyle.bold mixed/unavailable (null)");
  if (rowHeight === null) limitations.push("headerStyle.rowHeight mixed/unavailable (null)");
  return { fillColor, fontColor, bold, rowHeight };
}

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
      context.workbook.load("name");
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const items = sheets.items;
      if (!Array.isArray(items)) throw new Error("WorksheetCollection.items is not an array");
      if (items.length > WORKBOOK_TEMPLATE_MAX_SHEETS) {
        throw new Error(
          `resource-limit: workbook has ${items.length} sheets (max ${WORKBOOK_TEMPLATE_MAX_SHEETS})`,
        );
      }
      const workbookName = requireNonEmptyString(context.workbook.name, "Workbook.name");

      const captured: WorkbookTemplateCapturedSheet[] = [];
      const globalLimitations = [
        "Shallow snapshot only — not a full theme/table-style/CF/DV dump; not a replayable template package",
        "Not real Excel sideload verified",
      ];

      for (const sheet of items) {
        assertCaptureSheetSurface(sheet);
        sheet.load("name");
        const used = sheet.getUsedRangeOrNullObject(false) as ExcelRange & {
          isNullObject?: unknown;
        };
        await context.sync();
        const sheetName = requireNonEmptyString(sheet.name, "Worksheet.name");
        const isNull = requireBoolean(used.isNullObject, "UsedRange.isNullObject");
        const sheetLimitations: string[] = [];

        let usedRange: string | null = null;
        let rows = 0;
        let columns = 0;
        let baseStyle: WorkbookTemplateBaseStyle | null = null;
        let headerStyle: WorkbookTemplateHeaderStyle | null = null;

        if (isNull) {
          sheetLimitations.push("empty sheet: usedRange/baseStyle/headerStyle null");
        } else {
          used.load("address,rowCount,columnCount");
          await context.sync();
          const addr = requireNonEmptyString(used.address, "UsedRange.address");
          usedRange = splitSheetQualifiedAddress(addr).bare.replace(/\$/g, "").toUpperCase();
          rows = requirePositiveInt(used.rowCount, "UsedRange.rowCount");
          columns = requirePositiveInt(used.columnCount, "UsedRange.columnCount");
          baseStyle = await readBaseStyle(used, context, sheetLimitations);
          headerStyle = await readHeaderStyle(sheet, used, context, sheetLimitations);
        }

        const layout = await readSheetPageLayoutInContext(sheet, context);
        const print = printFromPageLayout(layout);

        captured.push({
          name: sheetName,
          usedRange,
          rows,
          columns,
          baseStyle,
          headerStyle,
          print,
          limitations: sheetLimitations,
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
        limitations: globalLimitations,
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
