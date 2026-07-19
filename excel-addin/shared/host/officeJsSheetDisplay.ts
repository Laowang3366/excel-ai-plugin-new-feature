import { withExcel } from "./officeJsRuntime";
import type {
  HostResult,
  SheetDisplayInfo,
  SheetDisplayUpdateInput,
} from "./types";

const DISPLAY_PROPS = "name,tabColor,showGridlines,showHeadings";

function toDisplayInfo(
  sheetName: string,
  sheet: { name: string; tabColor: string; showGridlines: boolean; showHeadings: boolean },
): SheetDisplayInfo {
  return {
    sheetName: sheet.name || sheetName,
    tabColor: sheet.tabColor ?? "",
    showGridlines: Boolean(sheet.showGridlines),
    showHeadings: Boolean(sheet.showHeadings),
  };
}

export async function officeJsGetSheetDisplay(
  sheetName: string,
): Promise<HostResult<SheetDisplayInfo>> {
  return withExcel("sheet.display.get", async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load(DISPLAY_PROPS);
    await context.sync();
    return toDisplayInfo(sheetName, sheet);
  });
}

export async function officeJsSetSheetDisplay(
  input: SheetDisplayUpdateInput,
): Promise<HostResult<SheetDisplayInfo>> {
  return withExcel("sheet.display.set", async (context) => {
    const sheet = context.workbook.worksheets.getItem(input.sheetName);
    if (input.tabColor != null) sheet.tabColor = input.tabColor;
    if (input.showGridlines != null) sheet.showGridlines = input.showGridlines;
    if (input.showHeadings != null) sheet.showHeadings = input.showHeadings;
    sheet.load(DISPLAY_PROPS);
    await context.sync();
    return toDisplayInfo(input.sheetName, sheet);
  });
}
