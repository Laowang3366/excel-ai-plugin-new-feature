import type { ExcelRequestContext, ExcelWorksheet } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type { HostResult, SheetFreezeInfo, SheetFreezeSetInput } from "./types";

function toFreezeInfo(
  sheetName: string,
  location: {
    isNullObject: boolean;
    address: string;
    rowCount: number;
    columnCount: number;
  },
): SheetFreezeInfo {
  if (location.isNullObject) {
    return { sheetName, address: null, rowCount: 0, columnCount: 0 };
  }
  return {
    sheetName,
    address: location.address,
    rowCount: location.rowCount,
    columnCount: location.columnCount,
  };
}

async function readLocation(
  sheet: ExcelWorksheet,
  context: ExcelRequestContext,
  sheetName: string,
): Promise<SheetFreezeInfo> {
  const location = sheet.freezePanes.getLocationOrNullObject();
  location.load("address,rowCount,columnCount");
  await context.sync();
  return toFreezeInfo(sheet.name || sheetName, location);
}

export async function officeJsGetSheetFreeze(
  sheetName: string,
): Promise<HostResult<SheetFreezeInfo>> {
  return withExcel("sheet.freeze.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    return readLocation(sheet, context, sheetName);
  });
}

export async function officeJsSetSheetFreeze(
  input: SheetFreezeSetInput,
): Promise<HostResult<SheetFreezeInfo>> {
  return withExcel("sheet.freeze.set", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    sheet.load("name");
    const panes = sheet.freezePanes;
    if (input.command === "rows") {
      panes.freezeRows(input.count!);
    } else if (input.command === "columns") {
      panes.freezeColumns(input.count!);
    } else if (input.command === "at") {
      panes.freezeAt(sheet.getRange(input.address!));
    } else {
      panes.unfreeze();
    }
    return readLocation(sheet, context, input.sheetName);
  });
}
