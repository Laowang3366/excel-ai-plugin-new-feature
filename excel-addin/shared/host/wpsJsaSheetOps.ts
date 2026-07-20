import { getSheet, requireWorkbook, type WpsSheet } from "./wpsJsaRuntime";
import type { HostResult, SheetInfo } from "./types";
import { fail, ok, unsupported } from "./types";

/**
 * Assumed Worksheet.Copy/Move (desktop ExcelObjectActionService COM parity).
 * Not in the in-repo JSA bridge contract; not device-verified.
 */
const EVIDENCE =
  "Assumed Worksheet.Copy/Move (desktop ExcelObjectActionService COM parity; not in bridge contract; not device-verified)";

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function toSheetInfo(sheet: WpsSheet, indexFallback: number, activeName?: string): SheetInfo {
  return {
    name: sheet.Name,
    index: sheet.Index ?? indexFallback,
    isActive: activeName != null ? sheet.Name === activeName : false,
  };
}

export async function wpsCopySheet(
  sheetName: string,
  newName?: string,
): Promise<HostResult<SheetInfo>> {
  const workbookResult = requireWorkbook("sheet.copy");
  if (!workbookResult.ok) return workbookResult;
  const workbook = workbookResult.data;
  const sheets = workbook.Worksheets;
  if (!sheets || typeof sheets.Count !== "number" || typeof sheets.Item !== "function") {
    return unsupported(
      "sheet.copy",
      "wps-jsa",
      "Worksheets collection unavailable",
      EVIDENCE,
    );
  }
  const sheet = getSheet(workbook, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.copy",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (typeof sheet.Copy !== "function") {
    return unsupported(
      "sheet.copy",
      "wps-jsa",
      "Worksheet.Copy is unavailable",
      EVIDENCE,
    );
  }

  try {
    const last = sheets.Item(sheets.Count);
    // Excel/WPS: Copy(Before, After) — desktop uses After: last sheet.
    sheet.Copy(undefined, last);
    const copied = workbook.ActiveSheet ?? sheets.Item(sheets.Count);
    if (!copied?.Name) {
      return fail(
        "sheet.copy",
        "wps-jsa",
        "Copy completed but ActiveSheet/new sheet is unavailable",
        EVIDENCE,
      );
    }
    if (newName != null && newName !== "") {
      copied.Name = newName;
    }
    return ok(toSheetInfo(copied, sheets.Count, workbook.ActiveSheet?.Name));
  } catch (error) {
    return fail("sheet.copy", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

/**
 * @param position 1-based public sheet position (desktop/COM parity)
 */
export async function wpsMoveSheet(
  sheetName: string,
  position: number,
): Promise<HostResult<SheetInfo>> {
  if (!Number.isInteger(position) || position < 1) {
    return fail(
      "sheet.move",
      "wps-jsa",
      "position must be a 1-based positive integer",
      EVIDENCE,
    );
  }

  const workbookResult = requireWorkbook("sheet.move");
  if (!workbookResult.ok) return workbookResult;
  const workbook = workbookResult.data;
  const sheets = workbook.Worksheets;
  if (!sheets || typeof sheets.Count !== "number" || typeof sheets.Item !== "function") {
    return unsupported(
      "sheet.move",
      "wps-jsa",
      "Worksheets collection unavailable",
      EVIDENCE,
    );
  }
  const sheet = getSheet(workbook, sheetName);
  if (!sheet) {
    return unsupported(
      "sheet.move",
      "wps-jsa",
      `Sheet "${sheetName}" not found`,
      EVIDENCE,
    );
  }
  if (typeof sheet.Move !== "function") {
    return unsupported(
      "sheet.move",
      "wps-jsa",
      "Worksheet.Move is unavailable",
      EVIDENCE,
    );
  }

  try {
    // Desktop ExcelObjectActionService: Move(Before: Item(Max(1, position))).
    const targetIndex = Math.max(1, position);
    const before =
      targetIndex > sheets.Count ? undefined : sheets.Item(Math.min(targetIndex, sheets.Count));
    if (before) {
      sheet.Move(before, undefined);
    } else {
      sheet.Move(undefined, sheets.Item(sheets.Count));
    }
    const moved = getSheet(workbook, sheet.Name) ?? sheet;
    return ok(toSheetInfo(moved, position, workbook.ActiveSheet?.Name));
  } catch (error) {
    return fail("sheet.move", "wps-jsa", messageOf(error), EVIDENCE);
  }
}
