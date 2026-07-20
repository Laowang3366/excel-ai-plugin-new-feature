/**
 * Office.js pivot.list (ExcelApi 1.8) + pivot.refresh (ExcelApi 1.3 / connections 1.7).
 */
import { queueLoadPivotHierarchies, pivotToInfo, withPivotExcel } from "./officeJsPivotShared";
import type {
  ExcelPivotTable,
  ExcelRequestContextWithPivot,
  ExcelWorksheetWithPivot,
} from "./officeJsPivotTypes";
import type {
  PivotConnectionRefreshInfo,
  PivotListInfo,
  PivotListInput,
  PivotRefreshInfo,
  PivotRefreshInput,
  PivotTableInfo,
} from "./pivotTypes";
import type { HostResult } from "./types";

const CONNECTION_LIMITATIONS = [
  "Office.js DataConnectionCollection.refreshAll only refreshes supported connections (e.g. PivotTable→Power BI dataset; same-workbook Data Model→table/range)",
  "Power Query connections are not supported by Workbook.dataConnections.refreshAll",
  "Data connections outside the original workbook are not supported (except Power BI)",
  "Firewall-protected data connections are not supported",
  "No DataConnection status/count/readback API; connection refresh is request-accepted only (verified:false), not full Workbook.RefreshAll parity",
] as const;

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

async function resolveRefreshTargets(
  context: ExcelRequestContextWithPivot,
  input: PivotRefreshInput,
): Promise<Array<{ pivot: ExcelPivotTable; sheetName: string }>> {
  const targets: Array<{ pivot: ExcelPivotTable; sheetName: string }> = [];

  if (input.sheetName && input.name) {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as ExcelWorksheetWithPivot;
    sheet.load("name");
    const pivot = sheet.pivotTables.getItem(input.name);
    pivot.load("name");
    await context.sync();
    targets.push({ pivot, sheetName: String(sheet.name) });
    return targets;
  }
  if (input.sheetName) {
    const sheet = context.workbook.worksheets.getItem(input.sheetName) as ExcelWorksheetWithPivot;
    sheet.load("name");
    sheet.pivotTables.load("items/name");
    await context.sync();
    for (const pivot of sheet.pivotTables.items ?? []) {
      targets.push({ pivot, sheetName: String(sheet.name) });
    }
    return targets;
  }
  if (input.name) {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    for (const sheet of sheets.items ?? []) {
      const ws = sheet as ExcelWorksheetWithPivot;
      ws.pivotTables.load("items/name");
    }
    await context.sync();
    let found: { pivot: ExcelPivotTable; sheetName: string } | null = null;
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
    return targets;
  }

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
  return targets;
}

function queueConnectionRefresh(context: ExcelRequestContextWithPivot): PivotConnectionRefreshInfo {
  const dc = context.workbook.dataConnections;
  if (!dc || typeof dc.refreshAll !== "function") {
    throw new Error("Workbook.dataConnections.refreshAll is not available");
  }
  dc.refreshAll();
  return {
    requested: true,
    method: "Workbook.dataConnections.refreshAll",
    verified: false,
    scope: "supported-office-js-connections",
  };
}

export async function officeJsRefreshPivots(
  input: PivotRefreshInput = {},
): Promise<HostResult<PivotRefreshInfo>> {
  const wantConnections = input.refreshConnections === true;
  // 1.7 gate when connections requested (includes 1.3 on real hosts); else 1.3 only.
  const version = wantConnections ? "1.7" : "1.3";

  return withPivotExcel(
    "pivot.refresh",
    async (context) => {
      // Resolve all pivot targets before any refresh side effects.
      const targets = await resolveRefreshTargets(context, input);

      if (targets.length === 0 && !wantConnections) {
        return { refreshed: [], count: 0, limitations: ["no pivot tables matched"] };
      }

      // Pre-check connection member before queuing pivot refresh (fail closed).
      if (wantConnections) {
        const dc = context.workbook.dataConnections;
        if (!dc || typeof dc.refreshAll !== "function") {
          throw new Error("Workbook.dataConnections.refreshAll is not available");
        }
      }

      for (const t of targets) {
        if (typeof t.pivot.refresh !== "function") {
          throw new Error("PivotTable.refresh is not available");
        }
        t.pivot.refresh();
      }

      let connectionRefresh: PivotConnectionRefreshInfo | undefined;
      const limitations: string[] = [];
      if (wantConnections) {
        connectionRefresh = queueConnectionRefresh(context);
        limitations.push(...CONNECTION_LIMITATIONS);
      }

      await context.sync();

      if (targets.length > 0) {
        for (const t of targets) t.pivot.load("name");
        await context.sync();
      }

      const refreshed = targets.map((t) => ({
        name: String(t.pivot.name),
        sheetName: t.sheetName,
        refreshed: true,
      }));
      return {
        refreshed,
        count: refreshed.length,
        ...(connectionRefresh ? { connectionRefresh } : {}),
        ...(limitations.length > 0 ? { limitations } : {}),
      };
    },
    version,
  );
}
