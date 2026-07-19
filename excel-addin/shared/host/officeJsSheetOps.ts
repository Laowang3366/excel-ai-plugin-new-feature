import type { ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type { HostResult, SheetInfo } from "./types";

export async function officeJsCopySheet(
  sheetName: string,
  newName?: string,
): Promise<HostResult<SheetInfo>> {
  return withExcel("sheet.copy", async (context: ExcelRequestContext) => {
    const source = context.workbook.worksheets.getItem(sheetName);
    source.load("name,position");
    await context.sync();
    const copied = source.copy("After", source);
    if (newName) copied.name = newName;
    copied.load("name,position");
    await context.sync();
    // Public index is 1-based (desktop sheet.operation position parity).
    return {
      name: copied.name,
      index: copied.position + 1,
      isActive: false,
    };
  });
}

/**
 * @param position1Based desktop/COM 1-based sheet position (minimum 1)
 */
export async function officeJsMoveSheet(
  sheetName: string,
  position1Based: number,
): Promise<HostResult<SheetInfo>> {
  return withExcel("sheet.move", async (context: ExcelRequestContext) => {
    if (!Number.isInteger(position1Based) || position1Based < 1) {
      throw new Error("position must be a 1-based positive integer");
    }
    const sheet = context.workbook.worksheets.getItem(sheetName);
    // Office.js Worksheet.position is 0-based.
    sheet.position = position1Based - 1;
    sheet.load("name,position");
    await context.sync();
    return {
      name: sheet.name,
      index: sheet.position + 1,
      isActive: false,
    };
  });
}
