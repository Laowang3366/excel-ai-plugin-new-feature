/**
 * workbook.template.apply — desktop applyWorkbookTemplate parity (Office.js).
 * write → sync → load → sync; strict readback; no request echo.
 */
import { parseA1Cell, toA1 } from "./a1Address";
import type { ExcelRange, ExcelRequestContext, ExcelWorksheet } from "./officeJsExcelTypes";
import { getExcelRun } from "./officeJsRuntime";
import { requireExcelApi18ForTemplateApply } from "./officeJsTemplateRequirements";
import {
  colorsEqual,
  numbersClose,
  requireAlignmentCenter,
  requireBoolean,
  requireFiniteNumber,
  requireHexColor,
  requireNonEmptyString,
  stripSheetPrefix,
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

function headerRowAddress(usedAddress: string, columnCount: number): string {
  const bare = stripSheetPrefix(usedAddress).replace(/\$/g, "");
  const start = bare.includes(":") ? bare.split(":")[0]! : bare;
  const cell = parseA1Cell(start);
  if (!cell) throw new Error(`Cannot parse usedRange address for header: ${usedAddress}`);
  const endCol = cell.col + Math.max(1, columnCount) - 1;
  return `${toA1(cell.row, cell.col)}:${toA1(cell.row, endCol)}`;
}

function assertSurface(sheet: ExcelWorksheet): void {
  if (typeof sheet.getUsedRangeOrNullObject !== "function") {
    throw new Error("Worksheet.getUsedRangeOrNullObject is missing");
  }
  if (typeof sheet.showGridlines !== "boolean" && sheet.showGridlines === undefined) {
    // may be unloaded; presence of property on prototype/object is enough pre-write check via 'in'
  }
  if (!("showGridlines" in sheet)) throw new Error("Worksheet.showGridlines is missing");
  if (!sheet.freezePanes || typeof sheet.freezePanes.unfreeze !== "function") {
    throw new Error("Worksheet.freezePanes is missing");
  }
  if (typeof sheet.freezePanes.freezeRows !== "function") {
    throw new Error("Worksheet.freezePanes.freezeRows is missing");
  }
  if (typeof sheet.getRange !== "function") throw new Error("Worksheet.getRange is missing");
}

function assertRangeFormatSurface(range: ExcelRange, needAutofit: boolean): void {
  if (!range.format) throw new Error("Range.format is missing");
  if (!range.format.font) throw new Error("Range.format.font is missing");
  if (!range.format.fill) throw new Error("Range.format.fill is missing");
  if (needAutofit) {
    if (typeof range.format.autofitColumns !== "function") {
      throw new Error("Range.format.autofitColumns is missing");
    }
    if (typeof range.format.autofitRows !== "function") {
      throw new Error("Range.format.autofitRows is missing");
    }
  }
}

async function isEmptyUsed(
  used: ExcelRange & { isNullObject: boolean },
  context: ExcelRequestContext,
): Promise<boolean> {
  if (used.isNullObject) return true;
  used.load("address,rowCount,columnCount,text");
  await context.sync();
  const rows = requireFiniteNumber(used.rowCount, "UsedRange.rowCount");
  const cols = requireFiniteNumber(used.columnCount, "UsedRange.columnCount");
  if (rows === 1 && cols === 1) {
    const text = (used as ExcelRange & { text?: unknown }).text;
    if (text == null || text === "") return true;
    if (typeof text === "string" && text.trim() === "") return true;
    // Prefer text; if host exposes values only for 1x1, load values once.
    if (typeof text !== "string") {
      used.load("values");
      await context.sync();
      const v = used.values?.[0]?.[0];
      if (v == null || (typeof v === "string" && v.trim() === "")) return true;
    }
  }
  return false;
}

async function resolveTargets(
  context: ExcelRequestContext,
  input: WorkbookTemplateApplyInput,
): Promise<ExcelWorksheet[]> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  const active = context.workbook.worksheets.getActiveWorksheet();
  active.load("name");
  await context.sync();

  const items = sheets.items;
  if (!Array.isArray(items)) throw new Error("WorksheetCollection.items is not an array");
  if (items.length > 500) {
    throw new Error("workbook exceeds 500 worksheets (resource-limit)");
  }

  const byLower = new Map<string, ExcelWorksheet>();
  for (const sheet of items) {
    const name = requireNonEmptyString(sheet.name, "Worksheet.name");
    byLower.set(name.toLowerCase(), sheet);
  }

  let selected: ExcelWorksheet[] = [];
  if (input.sheetNames && input.sheetNames.length > 0) {
    for (const requested of input.sheetNames) {
      const hit = byLower.get(requested.toLowerCase());
      if (!hit) throw new Error(`sheet not found: ${requested}`);
      selected.push(hit);
    }
  } else {
    selected = items.slice();
  }

  if (!input.allSheets) {
    const activeName = requireNonEmptyString(active.name, "ActiveWorksheet.name");
    selected = selected.filter(
      (s) => requireNonEmptyString(s.name, "Worksheet.name").toLowerCase() === activeName.toLowerCase(),
    );
  }

  // de-dupe preserving order
  const seen = new Set<string>();
  const unique: ExcelWorksheet[] = [];
  for (const sheet of selected) {
    const name = requireNonEmptyString(sheet.name, "Worksheet.name");
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(sheet);
  }
  if (unique.length === 0) {
    throw new Error("no target worksheets after sheetNames/allSheets selection");
  }
  return unique;
}

export async function officeJsApplyWorkbookTemplate(
  input: WorkbookTemplateApplyInput,
): Promise<HostResult<WorkbookTemplateApplyInfo>> {
  const gate = requireExcelApi18ForTemplateApply(CAPABILITY);
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
    const data = await run(async (context) => {
      const style = WORKBOOK_TEMPLATE_PRESET_STYLES[input.preset];
      const targets = await resolveTargets(context, input);
      for (const sheet of targets) assertSurface(sheet);

      type Plan =
        | { kind: "skip"; name: string; reason: string }
        | {
            kind: "apply";
            sheet: ExcelWorksheet;
            name: string;
            used: ExcelRange;
            header: ExcelRange;
          };

      const plans: Plan[] = [];
      for (const sheet of targets) {
        const name = requireNonEmptyString(sheet.name, "Worksheet.name");
        const used = sheet.getUsedRangeOrNullObject(false);
        await context.sync();
        if (await isEmptyUsed(used as ExcelRange & { isNullObject: boolean }, context)) {
          plans.push({ kind: "skip", name, reason: "empty used range (no non-blank content)" });
          continue;
        }
        const usedAddr = requireNonEmptyString(used.address, "UsedRange.address");
        const rows = requireFiniteNumber(used.rowCount, "UsedRange.rowCount");
        const cols = requireFiniteNumber(used.columnCount, "UsedRange.columnCount");
        assertRangeFormatSurface(used, input.autoFit);
        const header = sheet.getRange(headerRowAddress(usedAddr, cols));
        assertRangeFormatSurface(header, input.autoFit);
        plans.push({ kind: "apply", sheet, name, used, header });
        void rows;
      }

      // All validation done — perform writes
      for (const plan of plans) {
        if (plan.kind !== "apply") continue;
        const { sheet, used, header } = plan;
        used.format.font.name = input.fontName;
        used.format.font.size = input.fontSize;
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
        if (input.freezeRows > 0) {
          sheet.freezePanes.freezeRows(input.freezeRows);
        }
      }
      await context.sync();

      // Reload and verify
      const applied: WorkbookTemplateAppliedSheet[] = [];
      const skipped: WorkbookTemplateSkippedSheet[] = [];
      const limitations: string[] = [
        "autoFit does not verify exact columnWidth/rowHeight against a fixed size; autoFitVerified=false",
        "Not real Excel sideload verified",
      ];

      for (const plan of plans) {
        if (plan.kind === "skip") {
          skipped.push({ name: plan.name, reason: plan.reason });
          limitations.push(`skipped sheet ${plan.name}: ${plan.reason}`);
          continue;
        }
        const { sheet, used, header, name } = plan;
        used.load("address,rowCount,columnCount");
        used.format.font.load("name,size");
        header.format.font.load("name,size,bold,color");
        header.format.fill.load("color");
        header.format.load("horizontalAlignment,wrapText,rowHeight");
        sheet.load("name,showGridlines");
        const loc = sheet.freezePanes.getLocationOrNullObject();
        loc.load("address,rowCount,columnCount");
        await context.sync();

        const range = requireNonEmptyString(used.address, "UsedRange.address");
        const rows = requireFiniteNumber(used.rowCount, "UsedRange.rowCount");
        const columns = requireFiniteNumber(used.columnCount, "UsedRange.columnCount");
        const fontName = requireNonEmptyString(used.format.font.name, "font.name");
        const fontSize = requireFiniteNumber(used.format.font.size, "font.size");
        if (fontName !== input.fontName) {
          throw new Error(`fontName readback mismatch: ${fontName} !== ${input.fontName}`);
        }
        if (!numbersClose(fontSize, input.fontSize, 0.05)) {
          throw new Error(`fontSize readback mismatch: ${fontSize} !== ${input.fontSize}`);
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
        let freezeRowCount = 0;
        if (loc.isNullObject) {
          if (input.freezeRows !== 0) {
            throw new Error("freeze location is null but freezeRows > 0");
          }
        } else {
          freezeRowCount = requireFiniteNumber(loc.rowCount, "freeze.rowCount");
          if (freezeRowCount !== input.freezeRows) {
            throw new Error(
              `freeze rowCount readback mismatch: ${freezeRowCount} !== ${input.freezeRows}`,
            );
          }
        }

        applied.push({
          name: requireNonEmptyString(sheet.name, "Worksheet.name") || name,
          range: stripSheetPrefix(range).replace(/\$/g, ""),
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
