/**
 * Office.js pivot.list (ExcelApi 1.8) + pivot.refresh (ExcelApi 1.3).
 */
import { queueLoadPivotHierarchies, pivotToInfo, withPivotExcel } from "./officeJsPivotShared";
import type {
  ExcelPivotTable,
  ExcelWorksheetWithPivot,
} from "./officeJsPivotTypes";
import type {
  PivotListInfo,
  PivotListInput,
  PivotRefreshInfo,
  PivotRefreshInput,
  PivotTableInfo,
} from "./pivotTypes";
import type { HostResult } from "./types";
import { fail } from "./types";

export async function officeJsListPivots(
  input: PivotListInput = {},
): Promise<HostResult<PivotListInfo>> {
  return withPivotExcel("pivot.list", async (context) => {
    const limitations: string[] = [];
    const pivots: PivotTableInfo[] = [];

    if (input.sheetName) {
      const sheet = context.workbook.worksheets.getItem(input.sheetName) as ExcelWorksheetWithPivot;
      sheet.load("name");
      sheet.pivotTables.load("items/name");
      await context.sync();
      for (const pivot of sheet.pivotTables.items ?? []) {
        queueLoadPivotHierarchies(pivot);
      }
      await context.sync();
      for (const pivot of sheet.pivotTables.items ?? []) {
        pivots.push(pivotToInfo(pivot, String(sheet.name)));
      }
    } else {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        ws.pivotTables.load("items/name");
      }
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        for (const pivot of ws.pivotTables.items ?? []) {
          queueLoadPivotHierarchies(pivot);
        }
      }
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        for (const pivot of ws.pivotTables.items ?? []) {
          pivots.push(pivotToInfo(pivot, String(ws.name)));
        }
      }
    }

    pivots.sort((a, b) => {
      const sheetCmp = a.sheetName.localeCompare(b.sheetName);
      if (sheetCmp !== 0) return sheetCmp;
      return a.name.localeCompare(b.name);
    });
    limitations.push(
      "list uses Office.js hierarchy collections; Excel calculation-engine cycle analysis is not claimed",
    );
    return { pivots, limitations };
  });
}

export async function officeJsRefreshPivots(
  input: PivotRefreshInput = {},
): Promise<HostResult<PivotRefreshInfo>> {
  if (input.refreshConnections === true) {
    return fail(
      "pivot.refresh",
      "office-js",
      "refreshConnections is not supported on the add-in: desktop Workbook.RefreshAll has no proven Office.js equivalent (not desktop parity)",
      "Only PivotTable.refresh is implemented; external connections are not refreshed",
    );
  }

  return withPivotExcel("pivot.refresh", async (context) => {
    const targets: Array<{ pivot: ExcelPivotTable; sheetName: string }> = [];

    if (input.sheetName && input.name) {
      const sheet = context.workbook.worksheets.getItem(input.sheetName) as ExcelWorksheetWithPivot;
      sheet.load("name");
      const pivot = sheet.pivotTables.getItem(input.name);
      pivot.load("name");
      await context.sync();
      targets.push({ pivot, sheetName: String(sheet.name) });
    } else if (input.sheetName) {
      const sheet = context.workbook.worksheets.getItem(input.sheetName) as ExcelWorksheetWithPivot;
      sheet.load("name");
      sheet.pivotTables.load("items/name");
      await context.sync();
      for (const pivot of sheet.pivotTables.items ?? []) {
        targets.push({ pivot, sheetName: String(sheet.name) });
      }
    } else if (input.name) {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      let found: { pivot: ExcelPivotTable; sheetName: string } | null = null;
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        ws.pivotTables.load("items/name");
      }
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        for (const pivot of ws.pivotTables.items ?? []) {
          if (String(pivot.name).toLowerCase() === input.name!.toLowerCase()) {
            found = { pivot, sheetName: String(ws.name) };
            break;
          }
        }
        if (found) break;
      }
      if (!found) throw new Error(`pivot not found: ${input.name}`);
      targets.push(found);
    } else {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        ws.pivotTables.load("items/name");
      }
      await context.sync();
      for (const sheet of sheets.items ?? []) {
        const ws = sheet as ExcelWorksheetWithPivot;
        for (const pivot of ws.pivotTables.items ?? []) {
          targets.push({ pivot, sheetName: String(ws.name) });
        }
      }
    }

    if (targets.length === 0) {
      return { refreshed: [], count: 0, limitations: ["no pivot tables matched"] };
    }

    for (const t of targets) {
      if (typeof t.pivot.refresh !== "function") {
        throw new Error("PivotTable.refresh is not available");
      }
      t.pivot.refresh();
    }
    await context.sync();

    for (const t of targets) {
      t.pivot.load("name");
    }
    await context.sync();

    const refreshed = targets.map((t) => ({
      name: String(t.pivot.name),
      sheetName: t.sheetName,
      refreshed: true,
    }));
    return { refreshed, count: refreshed.length };
  }, "1.3");
}
