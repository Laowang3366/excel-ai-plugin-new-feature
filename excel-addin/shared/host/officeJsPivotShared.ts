/**
 * Shared Office.js pivot runner + read helpers.
 */
import {
  queueLoadPivotHierarchies,
  readDataHierarchySummaries,
  readHierarchyNames,
} from "./officeJsPivotFields";
import {
  requireExcelApiForPivot,
  type PivotExcelApiVersion,
} from "./officeJsPivotRequirements";
import type { ExcelPivotTable, ExcelRequestContextWithPivot } from "./officeJsPivotTypes";
import type { PivotTableInfo } from "./pivotTypes";
import { getExcelRun } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

export async function withPivotExcel<T>(
  capability: string,
  fn: (context: ExcelRequestContextWithPivot) => Promise<T>,
  version: PivotExcelApiVersion = "1.8",
): Promise<HostResult<T>> {
  const gate = requireExcelApiForPivot(capability, version);
  if (gate) return gate as HostResult<T>;
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      capability,
      "office-js",
      "Excel.run is not available in this runtime",
      "Requires Microsoft Office Excel with Office.js loaded",
    );
  }
  try {
    return ok(
      await run(
        fn as unknown as (ctx: import("./officeJsExcelTypes").ExcelRequestContext) => Promise<T>,
      ),
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(capability, "office-js", message);
  }
}

export function readSourceString(
  pivot: ExcelPivotTable,
): { source: string | null; limitation?: string } {
  if (typeof pivot.getDataSourceString !== "function") {
    return {
      source: null,
      limitation: "PivotTable.getDataSourceString is unavailable; source not reported",
    };
  }
  try {
    const value = pivot.getDataSourceString();
    return { source: value == null || value === "" ? null : String(value) };
  } catch (error) {
    return {
      source: null,
      limitation: `getDataSourceString failed: ${error instanceof Error ? error.message : String(error)}`,
    };
  }
}

export function pivotToInfo(pivot: ExcelPivotTable, sheetName: string): PivotTableInfo {
  const limitations: string[] = [];
  const { source, limitation } = readSourceString(pivot);
  if (limitation) limitations.push(limitation);
  let destination: string | null = null;
  try {
    const address = pivot.layout.getRange().address;
    destination = address == null || address === "" ? null : String(address);
  } catch {
    limitations.push("layout.getRange address unavailable");
  }
  if (destination == null) limitations.push("destination range not reliably readable");
  return {
    name: String(pivot.name),
    sheetName: pivot.worksheet?.name ? String(pivot.worksheet.name) : sheetName,
    source,
    destination,
    rowFields: readHierarchyNames(pivot.rowHierarchies),
    columnFields: readHierarchyNames(pivot.columnHierarchies),
    filterFields: readHierarchyNames(pivot.filterHierarchies),
    dataFields: readDataHierarchySummaries(pivot.dataHierarchies),
    refreshed: null,
    limitations: limitations.length > 0 ? limitations : undefined,
  };
}

export { queueLoadPivotHierarchies };
