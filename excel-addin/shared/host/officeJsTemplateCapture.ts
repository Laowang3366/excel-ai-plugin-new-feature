/**
 * workbook.template.capture — shallow desktop capture/inspect parity (Office.js ExcelApi 1.9).
 */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";
import { officeJsGetSheetPageLayout } from "./officeJsPageLayout";
import { getExcelRun } from "./officeJsRuntime";
import { requireExcelApi19ForTemplateCapture } from "./officeJsTemplateRequirements";
import {
  optionalBooleanOrNull,
  optionalFiniteOrNull,
  optionalStringOrNull,
  requireFiniteNumber,
  requireHexColor,
  requireNonEmptyString,
  stripSheetPrefix,
} from "./officeJsTemplateReadback";
import type {
  WorkbookTemplateBaseStyle,
  WorkbookTemplateCaptureInfo,
  WorkbookTemplateCapturedSheet,
  WorkbookTemplateHeaderStyle,
  WorkbookTemplatePrintSnapshot,
} from "./workbookTemplateTypes";
import { WORKBOOK_TEMPLATE_MAX_SHEETS } from "./workbookTemplateTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const CAPABILITY = "workbook.template.capture";

async function readBaseStyle(
  used: ExcelRange,
  context: ExcelRequestContext,
  limitations: string[],
): Promise<WorkbookTemplateBaseStyle> {
  used.format.font.load("name,size,color");
  await context.sync();
  let fontName: string | null = null;
  let fontSize: number | null = null;
  let fontColor: string | null = null;
  try {
    fontName = optionalStringOrNull(used.format.font.name, "base.font.name");
    if (fontName !== null && fontName.trim() === "") fontName = null;
  } catch {
    fontName = null;
    limitations.push("baseStyle.fontName unavailable or mixed");
  }
  try {
    fontSize = optionalFiniteOrNull(used.format.font.size, "base.font.size");
  } catch {
    fontSize = null;
    limitations.push("baseStyle.fontSize unavailable or mixed");
  }
  try {
    if (used.format.font.color == null) fontColor = null;
    else fontColor = requireHexColor(used.format.font.color, "base.font.color");
  } catch {
    fontColor = null;
    limitations.push("baseStyle.fontColor unavailable or mixed");
  }
  return { fontName, fontSize, fontColor };
}

async function readHeaderStyle(
  sheet: ExcelWorksheet,
  used: ExcelRange,
  context: ExcelRequestContext,
  limitations: string[],
): Promise<WorkbookTemplateHeaderStyle> {
  const addr = requireNonEmptyString(used.address, "UsedRange.address");
  const bare = stripSheetPrefix(addr).replace(/\$/g, "");
  const start = bare.includes(":") ? bare.split(":")[0]! : bare;
  used.load("columnCount");
  await context.sync();
  const cols = requireFiniteNumber(used.columnCount, "UsedRange.columnCount");
  const cell = parseA1Cell(start);
  if (!cell) throw new Error(`Cannot parse usedRange for header: ${addr}`);
  const headerAddr = `${toA1(cell.row, cell.col)}:${toA1(cell.row, cell.col + cols - 1)}`;
  const header = sheet.getRange(headerAddr);
  header.format.font.load("bold,color");
  header.format.fill.load("color");
  header.format.load("rowHeight");
  await context.sync();

  let fillColor: string | null = null;
  let fontColor: string | null = null;
  let bold: boolean | null = null;
  let rowHeight: number | null = null;
  try {
    fillColor = requireHexColor(header.format.fill.color, "header.fill.color");
  } catch {
    fillColor = null;
    limitations.push("headerStyle.fillColor unavailable or mixed");
  }
  try {
    fontColor = requireHexColor(header.format.font.color, "header.font.color");
  } catch {
    fontColor = null;
    limitations.push("headerStyle.fontColor unavailable or mixed");
  }
  try {
    bold = optionalBooleanOrNull(header.format.font.bold, "header.font.bold");
  } catch {
    bold = null;
    limitations.push("headerStyle.bold unavailable or mixed");
  }
  try {
    rowHeight = optionalFiniteOrNull(header.format.rowHeight, "header.rowHeight");
  } catch {
    rowHeight = null;
    limitations.push("headerStyle.rowHeight unavailable or mixed");
  }
  return { fillColor, fontColor, bold, rowHeight };
}

function printFromPageLayout(layout: {
  printArea: string | null;
  orientation: string;
  paperSize: string;
  fitToPagesWide: number | null;
  fitToPagesTall: number | null;
  printTitleRows: string | null;
  printTitleColumns: string | null;
  headers: { center: string };
  footers: { center: string };
}): WorkbookTemplatePrintSnapshot {
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
    const skeleton = await run(async (context) => {
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
      const sheetNames: string[] = [];
      for (const sheet of items) {
        sheetNames.push(requireNonEmptyString(sheet.name, "Worksheet.name"));
      }
      return { workbookName, sheetNames };
    });

    const captured: WorkbookTemplateCapturedSheet[] = [];
    const globalLimitations = [
      "Shallow snapshot only — not a full theme/table-style/CF/DV dump; not a replayable template package",
      "Not real Excel sideload verified",
    ];

    for (const sheetName of skeleton.sheetNames) {
      const sheetLimitations: string[] = [];
      const usedMeta = await run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.load("name");
        if (typeof sheet.getUsedRangeOrNullObject !== "function") {
          throw new Error("Worksheet.getUsedRangeOrNullObject is missing");
        }
        const used = sheet.getUsedRangeOrNullObject(false);
        await context.sync();
        if (used.isNullObject) {
          return { empty: true as const, name: requireNonEmptyString(sheet.name, "Worksheet.name") };
        }
        used.load("address,rowCount,columnCount");
        await context.sync();
        return {
          empty: false as const,
          name: requireNonEmptyString(sheet.name, "Worksheet.name"),
          address: requireNonEmptyString(used.address, "UsedRange.address"),
          rows: requireFiniteNumber(used.rowCount, "UsedRange.rowCount"),
          columns: requireFiniteNumber(used.columnCount, "UsedRange.columnCount"),
        };
      });

      let baseStyle: WorkbookTemplateBaseStyle | null = null;
      let headerStyle: WorkbookTemplateHeaderStyle | null = null;
      let usedRange: string | null = null;
      let rows = 0;
      let columns = 0;

      if (usedMeta.empty) {
        sheetLimitations.push("empty sheet: usedRange/baseStyle/headerStyle null");
      } else {
        usedRange = stripSheetPrefix(usedMeta.address).replace(/\$/g, "");
        rows = usedMeta.rows;
        columns = usedMeta.columns;
        const styles = await run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(sheetName);
          const used = sheet.getUsedRangeOrNullObject(false);
          await context.sync();
          if (used.isNullObject) throw new Error("used range became empty during capture");
          const base = await readBaseStyle(used, context, sheetLimitations);
          const header = await readHeaderStyle(sheet, used, context, sheetLimitations);
          return { base, header };
        });
        baseStyle = styles.base;
        headerStyle = styles.header;
      }

      const layoutResult = await officeJsGetSheetPageLayout(sheetName);
      if (!layoutResult.ok) {
        throw new Error(layoutResult.reason ?? `pageLayout read failed for sheet ${sheetName}`);
      }
      const print = printFromPageLayout(layoutResult.data);

      captured.push({
        name: usedMeta.name,
        usedRange,
        rows,
        columns,
        baseStyle,
        headerStyle,
        print,
        limitations: sheetLimitations,
      });
    }

    return ok({
      template: {
        version: 1,
        capturedFrom: skeleton.workbookName,
        capturedAt: new Date().toISOString(),
        sheets: captured,
      },
      sheetCount: captured.length,
      limitations: globalLimitations,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.startsWith("resource-limit:")) {
      return fail(CAPABILITY, "office-js", message, "max 500 worksheets");
    }
    return fail(CAPABILITY, "office-js", message);
  }
}
