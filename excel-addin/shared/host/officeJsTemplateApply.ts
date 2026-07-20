/** workbook.template.apply — plan → precheck → write → sync → load → sync readback. */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelRange, ExcelWorksheet } from "./officeJsExcelTypes";
import { getExcelRun } from "./officeJsRuntime";
import {
  isEmptyUsedRange,
  resolveApplyTargets,
  verifyFreezeReadback,
} from "./officeJsTemplateApplyHelpers";
import { requireExcelApi18ForTemplateApply } from "./officeJsTemplateRequirements";
import {
  colorsEqual,
  fontsEqual,
  normalizeRangeAddressForCompare,
  numbersClose,
  requireAlignmentCenter,
  requireBoolean,
  requireFiniteNumber,
  requireHexColor,
  requireNonEmptyString,
  requirePositiveInt,
  splitSheetQualifiedAddress,
} from "./officeJsTemplateReadback";
import type {
  WorkbookTemplateAppliedSheet,
  WorkbookTemplateApplyInfo,
  WorkbookTemplateApplyInput,
  WorkbookTemplateSkippedSheet,
} from "./workbookTemplateTypes";
import {
  WORKBOOK_TEMPLATE_HEADER_ROW_HEIGHT,
  WORKBOOK_TEMPLATE_PRESET_STYLES,
} from "./workbookTemplateTypes";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

const CAPABILITY = "workbook.template.apply";

type ApplyPlan =
  | { kind: "skip"; name: string; reason: string }
  | {
      kind: "apply";
      name: string;
      sheet: ExcelWorksheet;
      used: ExcelRange;
      header: ExcelRange;
      expectedAddress: string;
      expectedRows: number;
      expectedCols: number;
    };

function headerRowAddress(usedAddress: string, columnCount: number): string {
  const bare = splitSheetQualifiedAddress(usedAddress).bare;
  const start = bare.includes(":") ? bare.split(":")[0]! : bare;
  const cell = parseA1Cell(start);
  if (!cell) throw new Error(`Cannot parse usedRange address for header: ${usedAddress}`);
  const endCol = cell.col + Math.max(1, columnCount) - 1;
  return `${toA1(cell.row, cell.col)}:${toA1(cell.row, endCol)}`;
}

function assertSheetSurface(sheet: ExcelWorksheet): void {
  if (typeof sheet.getUsedRangeOrNullObject !== "function") {
    throw new Error("Worksheet.getUsedRangeOrNullObject is missing");
  }
  if (!("showGridlines" in sheet)) throw new Error("Worksheet.showGridlines is missing");
  if (!sheet.freezePanes) throw new Error("Worksheet.freezePanes is missing");
  if (typeof sheet.freezePanes.unfreeze !== "function") {
    throw new Error("Worksheet.freezePanes.unfreeze is missing");
  }
  if (typeof sheet.freezePanes.freezeRows !== "function") {
    throw new Error("Worksheet.freezePanes.freezeRows is missing");
  }
  if (typeof sheet.freezePanes.getLocationOrNullObject !== "function") {
    throw new Error("Worksheet.freezePanes.getLocationOrNullObject is missing");
  }
  if (typeof sheet.getRange !== "function") throw new Error("Worksheet.getRange is missing");
  if (typeof sheet.load !== "function") throw new Error("Worksheet.load is missing");
}

function assertRangeWriteSurface(
  range: ExcelRange,
  needAutofit: boolean,
  label: string,
): void {
  if (typeof range.load !== "function") throw new Error(`${label}.load is missing`);
  if (!range.format) throw new Error(`${label}.format is missing`);
  if (typeof range.format.load !== "function") {
    throw new Error(`${label}.format.load is missing`);
  }
  if (!range.format.font) throw new Error(`${label}.format.font is missing`);
  if (typeof range.format.font.load !== "function") {
    throw new Error(`${label}.format.font.load is missing`);
  }
  if (!range.format.fill) throw new Error(`${label}.format.fill is missing`);
  if (typeof range.format.fill.load !== "function") {
    throw new Error(`${label}.format.fill.load is missing`);
  }
  if (needAutofit) {
    if (typeof range.format.autofitColumns !== "function") {
      throw new Error(`${label}.format.autofitColumns is missing`);
    }
    if (typeof range.format.autofitRows !== "function") {
      throw new Error(`${label}.format.autofitRows is missing`);
    }
  }
}

export async function officeJsApplyWorkbookTemplate(
  input: WorkbookTemplateApplyInput,
): Promise<HostResult<WorkbookTemplateApplyInfo>> {
  const gate = requireExcelApi18ForTemplateApply(CAPABILITY);
  if (gate) return gate;

  if (input.sheetNames !== undefined && input.sheetNames.length === 0) {
    return fail(CAPABILITY, "office-js", "sheetNames must not be an empty array");
  }

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
    const data = await run(async (context) => {
      const style = WORKBOOK_TEMPLATE_PRESET_STYLES[input.preset];
      const targets = await resolveApplyTargets(context, input);
      for (const sheet of targets) assertSheetSurface(sheet);

      const plans: ApplyPlan[] = [];
      for (const sheet of targets) {
        const name = requireNonEmptyString(sheet.name, "Worksheet.name");
        const used = sheet.getUsedRangeOrNullObject(false) as ExcelRange & {
          isNullObject?: unknown;
        };
        await context.sync();
        if (await isEmptyUsedRange(used, context)) {
          plans.push({
            kind: "skip",
            name,
            reason: "empty used range (no non-blank content)",
          });
          continue;
        }
        const usedAddr = requireNonEmptyString(used.address, "UsedRange.address");
        const rows = requirePositiveInt(used.rowCount, "UsedRange.rowCount");
        const cols = requirePositiveInt(used.columnCount, "UsedRange.columnCount");
        const expectedAddress = normalizeRangeAddressForCompare(usedAddr);
        const header = sheet.getRange(headerRowAddress(usedAddr, cols));
        assertRangeWriteSurface(used, input.autoFit, "UsedRange");
        assertRangeWriteSurface(header, false, "HeaderRange");
        plans.push({
          kind: "apply",
          name,
          sheet,
          used,
          header,
          expectedAddress,
          expectedRows: rows,
          expectedCols: cols,
        });
      }

      const freezeProbes: Array<ExcelRange & { isNullObject?: unknown }> = [];
      for (const plan of plans) {
        if (plan.kind !== "apply") continue;
        const probe = plan.sheet.freezePanes.getLocationOrNullObject() as ExcelRange & {
          isNullObject?: unknown;
        };
        if (!probe || typeof probe.load !== "function") {
          throw new Error("freeze getLocationOrNullObject surface incomplete");
        }
        probe.load("isNullObject,address,rowCount,columnCount");
        freezeProbes.push(probe);
      }
      if (freezeProbes.length > 0) await context.sync();

      for (const plan of plans) {
        if (plan.kind !== "apply") continue;
        const { sheet, used, header } = plan;
        used.format.font.name = input.fontName;
        used.format.font.size = input.fontSize;
        header.format.font.name = input.fontName;
        header.format.font.size = input.fontSize;
        header.format.font.bold = true;
        header.format.font.color = style.headerFontColor;
        header.format.fill.color = style.headerFill;
        header.format.horizontalAlignment = "Center";
        header.format.wrapText = true;
        header.format.rowHeight = WORKBOOK_TEMPLATE_HEADER_ROW_HEIGHT;
        if (input.autoFit) {
          used.format.autofitColumns();
          used.format.autofitRows();
          header.format.rowHeight = WORKBOOK_TEMPLATE_HEADER_ROW_HEIGHT;
        }
        sheet.showGridlines = input.showGridlines;
        sheet.freezePanes.unfreeze();
        if (input.freezeRows > 0) sheet.freezePanes.freezeRows(input.freezeRows);
      }
      await context.sync();

      const applied: WorkbookTemplateAppliedSheet[] = [];
      const skipped: WorkbookTemplateSkippedSheet[] = [];
      const limitations: string[] = [
        "autoFit does not verify exact columnWidth/rowHeight against a fixed size; autoFitVerified=false",
        "Multi-sheet apply has no add-in-level rollback",
        "Not real Excel sideload verified",
      ];

      for (const plan of plans) {
        if (plan.kind === "skip") {
          skipped.push({ name: plan.name, reason: plan.reason });
          limitations.push(`skipped sheet ${plan.name}: ${plan.reason}`);
          continue;
        }
        const { sheet, used, header, name, expectedAddress, expectedRows, expectedCols } = plan;
        used.load("address,rowCount,columnCount");
        used.format.font.load("name,size");
        header.format.font.load("name,size,bold,color");
        header.format.fill.load("color");
        header.format.load("horizontalAlignment,wrapText,rowHeight");
        sheet.load("name,showGridlines");
        const loc = sheet.freezePanes.getLocationOrNullObject() as ExcelRange & {
          isNullObject?: unknown;
        };
        loc.load("isNullObject,address,rowCount,columnCount");
        await context.sync();

        const rangeRaw = requireNonEmptyString(used.address, "UsedRange.address");
        const rangeNorm = normalizeRangeAddressForCompare(rangeRaw);
        if (rangeNorm !== expectedAddress) {
          throw new Error(
            `usedRange address readback mismatch: ${rangeNorm} !== ${expectedAddress}`,
          );
        }
        const rows = requirePositiveInt(used.rowCount, "UsedRange.rowCount");
        const columns = requirePositiveInt(used.columnCount, "UsedRange.columnCount");
        if (rows !== expectedRows) {
          throw new Error(`usedRange rowCount mismatch: ${rows} !== ${expectedRows}`);
        }
        if (columns !== expectedCols) {
          throw new Error(`usedRange columnCount mismatch: ${columns} !== ${expectedCols}`);
        }

        const fontName = requireNonEmptyString(used.format.font.name, "font.name");
        const fontSize = requireFiniteNumber(used.format.font.size, "font.size");
        if (!fontsEqual(fontName, input.fontName)) {
          throw new Error(`fontName readback mismatch: ${fontName} !== ${input.fontName}`);
        }
        if (!numbersClose(fontSize, input.fontSize, 0.05)) {
          throw new Error(`fontSize readback mismatch: ${fontSize} !== ${input.fontSize}`);
        }
        const headerFontName = requireNonEmptyString(
          header.format.font.name,
          "header.font.name",
        );
        const headerFontSize = requireFiniteNumber(header.format.font.size, "header.font.size");
        if (!fontsEqual(headerFontName, input.fontName)) {
          throw new Error(
            `header fontName readback mismatch: ${headerFontName} !== ${input.fontName}`,
          );
        }
        if (!numbersClose(headerFontSize, input.fontSize, 0.05)) {
          throw new Error(
            `header fontSize readback mismatch: ${headerFontSize} !== ${input.fontSize}`,
          );
        }
        const headerFill = requireHexColor(header.format.fill.color, "header.fill.color");
        const headerFontColor = requireHexColor(header.format.font.color, "header.font.color");
        if (!colorsEqual(headerFill, style.headerFill)) {
          throw new Error(`header fill readback mismatch: ${headerFill}`);
        }
        if (!colorsEqual(headerFontColor, style.headerFontColor)) {
          throw new Error(`header font color readback mismatch: ${headerFontColor}`);
        }
        const headerBold = requireBoolean(header.format.font.bold, "header.font.bold");
        if (headerBold !== true) throw new Error("header bold readback is not true");
        const headerHorizontalAlignment = requireAlignmentCenter(
          header.format.horizontalAlignment,
          "header.horizontalAlignment",
        );
        const headerWrapText = requireBoolean(header.format.wrapText, "header.wrapText");
        if (headerWrapText !== true) throw new Error("header wrapText readback is not true");
        const headerRowHeight = requireFiniteNumber(header.format.rowHeight, "header.rowHeight");
        if (!numbersClose(headerRowHeight, WORKBOOK_TEMPLATE_HEADER_ROW_HEIGHT, 0.51)) {
          throw new Error(`header rowHeight readback mismatch: ${headerRowHeight}`);
        }
        const showGridlines = requireBoolean(sheet.showGridlines, "showGridlines");
        if (showGridlines !== input.showGridlines) {
          throw new Error(`showGridlines readback mismatch: ${showGridlines}`);
        }
        const freezeRowCount = verifyFreezeReadback(loc, input.freezeRows);

        applied.push({
          name: requireNonEmptyString(sheet.name, "Worksheet.name") || name,
          range: rangeNorm,
          rows,
          columns,
          readback: {
            fontName,
            fontSize,
            headerFill: headerFill.toUpperCase(),
            headerFontColor: headerFontColor.toUpperCase(),
            headerBold,
            headerHorizontalAlignment,
            headerWrapText,
            headerRowHeight,
            showGridlines,
            freezeRowCount,
            autoFitVerified: false,
          },
        });
      }

      return {
        preset: input.preset,
        appliedSheets: applied,
        appliedSheetCount: applied.length,
        skippedSheets: skipped,
        limitations,
      } satisfies WorkbookTemplateApplyInfo;
    });
    return ok(data);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(CAPABILITY, "office-js", message);
  }
}
