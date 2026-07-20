/**
 * Office.js pivot.create (ExcelApi 1.8) + re-exports list/refresh.
 */
import { validateBareA1 } from "./officeJsChartSource";
import {
  applyPivotFieldPlan,
  buildPivotFieldPlan,
  queueLoadPivotHierarchies,
  readDataHierarchySummaries,
  readHierarchyNames,
  type PivotFieldPlan,
} from "./officeJsPivotFields";
import {
  ensurePivotSheet,
  formatSheetA1,
  nextPivotDestinationAddress,
  parsePivotDestination,
} from "./officeJsPivotDestination";
import type {
  ExcelPivotTable,
  ExcelWorksheetWithPivot,
} from "./officeJsPivotTypes";
import type {
  PivotCreateInfo,
  PivotCreateInput,
  PivotCreateVerification,
} from "./pivotTypes";
import { readSourceString, withPivotExcel } from "./officeJsPivotShared";
import type { HostResult } from "./types";
import { fail } from "./types";

export { officeJsListPivots, officeJsRefreshPivots } from "./officeJsPivotQuery";

function buildVerification(
  pivot: ExcelPivotTable,
  expectedName: string,
  plan: PivotFieldPlan,
  destinationAddress: string | null,
): PivotCreateVerification {
  const checks: PivotCreateVerification["checks"] = [];
  const nameMatches = String(pivot.name) === expectedName;
  checks.push({ name: "name", ok: nameMatches, message: String(pivot.name) });

  const rowFieldCount = pivot.rowHierarchies.items?.length ?? 0;
  const columnFieldCount = pivot.columnHierarchies.items?.length ?? 0;
  const filterFieldCount = pivot.filterHierarchies.items?.length ?? 0;
  const dataFieldCount = pivot.dataHierarchies.items?.length ?? 0;

  const rowOk = rowFieldCount === plan.rowFields.length;
  const colOk = columnFieldCount === plan.columnFields.length;
  const filterOk = filterFieldCount === plan.filterFields.length;
  const dataOk = dataFieldCount === plan.dataFields.length;
  checks.push({ name: "rowFieldCount", ok: rowOk, message: String(rowFieldCount) });
  checks.push({ name: "columnFieldCount", ok: colOk, message: String(columnFieldCount) });
  checks.push({ name: "filterFieldCount", ok: filterOk, message: String(filterFieldCount) });
  checks.push({ name: "dataFieldCount", ok: dataOk, message: String(dataFieldCount) });

  const destinationReadable = destinationAddress != null && destinationAddress !== "";
  checks.push({
    name: "destinationReadable",
    ok: destinationReadable,
    message: destinationAddress ?? "(empty)",
  });

  const totalFields =
    rowFieldCount + columnFieldCount + filterFieldCount + dataFieldCount;
  const hasFields = totalFields > 0;
  checks.push({
    name: "hasFields",
    ok: hasFields,
    message: String(totalFields),
  });

  const okAll =
    nameMatches &&
    rowOk &&
    colOk &&
    filterOk &&
    dataOk &&
    destinationReadable &&
    hasFields;
  return {
    ok: okAll,
    objectExists: true,
    nameMatches,
    destinationReadable,
    rowFieldCount,
    columnFieldCount,
    filterFieldCount,
    dataFieldCount,
    checks,
  };
}

export async function officeJsCreatePivot(
  input: PivotCreateInput,
): Promise<HostResult<PivotCreateInfo>> {
  let plan: PivotFieldPlan;
  try {
    plan = buildPivotFieldPlan(input);
  } catch (error) {
    return fail(
      "pivot.create",
      "office-js",
      error instanceof Error ? error.message : String(error),
    );
  }

  let sourceAddress: string;
  try {
    sourceAddress = validateBareA1(input.sourceAddress, "sourceAddress");
  } catch (error) {
    return fail(
      "pivot.create",
      "office-js",
      error instanceof Error ? error.message : String(error),
    );
  }

  let destPlan;
  try {
    destPlan = parsePivotDestination(input.destination);
  } catch (error) {
    return fail(
      "pivot.create",
      "office-js",
      error instanceof Error ? error.message : String(error),
    );
  }

  const pivotName =
    input.name?.trim() ||
    `AI_Pivot_${new Date().toISOString().replace(/[-:TZ.]/g, "").slice(0, 17)}`;

  return withPivotExcel("pivot.create", async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(
      input.sourceSheetName,
    ) as ExcelWorksheetWithPivot;
    sourceSheet.load("name");
    await context.sync();

    let destSheet: ExcelWorksheetWithPivot;
    let destBare: string;
    if (destPlan.useDedicatedSheet) {
      destSheet = await ensurePivotSheet(context);
      destBare = await nextPivotDestinationAddress(context, destSheet);
    } else {
      const destSheetName = destPlan.sheetName ?? String(sourceSheet.name);
      destSheet = context.workbook.worksheets.getItem(destSheetName) as ExcelWorksheetWithPivot;
      destSheet.load("name");
      await context.sync();
      destBare = destPlan.address;
    }

    const sourceRange = sourceSheet.getRange(sourceAddress);
    const destRange = destSheet.getRange(destBare);
    sourceRange.load("address");
    destRange.load("address");
    await context.sync();

    if (typeof destSheet.pivotTables?.add !== "function") {
      throw new Error("Worksheet.pivotTables.add is not available");
    }

    const pivot = destSheet.pivotTables.add(pivotName, sourceRange, destRange);
    applyPivotFieldPlan(pivot, plan);
    queueLoadPivotHierarchies(pivot);
    await context.sync();

    queueLoadPivotHierarchies(pivot);
    await context.sync();

    let destinationAddress: string | null = null;
    try {
      destinationAddress = String(pivot.layout.getRange().address ?? "");
    } catch {
      destinationAddress = formatSheetA1(String(destSheet.name), destBare);
    }
    if (!destinationAddress) {
      destinationAddress = formatSheetA1(String(destSheet.name), destBare);
    }

    const verification = buildVerification(pivot, pivotName, plan, destinationAddress);
    if (!verification.ok) {
      throw new Error(
        `pivot_verification_failed: ${verification.checks
          .filter((c) => !c.ok)
          .map((c) => c.name)
          .join(",")}`,
      );
    }

    const { source, limitation } = readSourceString(pivot);
    const limitations: string[] = [];
    if (limitation) limitations.push(limitation);
    if (source == null) {
      limitations.push("source string not available after create; reported source is request address");
    }

    const sheetName = String(destSheet.name);
    return {
      name: String(pivot.name),
      sheetName,
      source: source ?? formatSheetA1(String(sourceSheet.name), sourceAddress),
      destination: destinationAddress,
      rowFields: readHierarchyNames(pivot.rowHierarchies),
      columnFields: readHierarchyNames(pivot.columnHierarchies),
      filterFields: readHierarchyNames(pivot.filterHierarchies),
      dataFields: readDataHierarchySummaries(pivot.dataHierarchies),
      verification,
      limitations: limitations.length > 0 ? limitations : undefined,
    };
  });
}
