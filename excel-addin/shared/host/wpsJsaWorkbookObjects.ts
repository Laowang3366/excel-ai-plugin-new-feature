/**
 * workbook.objects.inspect — WPS JSA partial inventory.
 * Sheets + named ranges via member-probe paths already implemented.
 * table/chart/shape remain typed unsupported categories (partial success).
 */
import { requireWorkbook } from "./wpsJsaRuntime";
import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import { fail, ok } from "./types";
import {
  availableCategory,
  failedCategory,
  sanitizeInventoryMessage,
  sortNamedRanges,
  sortSheets,
  unsupportedCategory,
} from "./workbookObjectsHelpers";
import type {
  WorkbookObjectsInspectInfo,
  WorkbookObjectsInspectInput,
} from "./workbookObjectsTypes";
import { WORKBOOK_OBJECTS_MAX_DEFAULT } from "./workbookObjectsTypes";

const TABLE_EVIDENCE = "No in-repo WPS ListObjects contract";
const CHART_EVIDENCE = "No in-repo WPS ChartObjects contract";
const SHAPE_EVIDENCE = "No in-repo WPS Shapes/ShapeObjects contract";

export async function wpsInspectWorkbookObjects(
  input: WorkbookObjectsInspectInput,
  deps: {
    listSheets: () => Promise<HostResult<SheetInfo[]>>;
    listNamedRanges: (input?: {
      scope?: "workbook" | "worksheet";
      sheetName?: string;
    }) => Promise<HostResult<NamedRangeInfo[]>>;
  },
): Promise<HostResult<WorkbookObjectsInspectInfo>> {
  const maxItems = input.maxItemsPerCategory ?? WORKBOOK_OBJECTS_MAX_DEFAULT;
  const filterSheetName = input.sheetName?.trim() || undefined;
  const limitations: string[] = [
    "WPS inventory uses member-probe paths only; not official JSA contract; not real-device verified.",
    "table/chart/shape categories are typed unsupported on WPS (not empty success).",
  ];

  const workbookResult = requireWorkbook("workbook.objects.inspect");
  if (!workbookResult.ok) return workbookResult;
  const workbook = workbookResult.data;

  const sheetsResult = await deps.listSheets();
  if (!sheetsResult.ok) {
    if (sheetsResult.unsupported) {
      return fail(
        "workbook.objects.inspect",
        "wps-jsa",
        sanitizeInventoryMessage(sheetsResult.reason),
        sheetsResult.evidence,
      );
    }
    return fail(
      "workbook.objects.inspect",
      "wps-jsa",
      sanitizeInventoryMessage(sheetsResult.reason),
      sheetsResult.evidence,
    );
  }

  const allSheets = sortSheets(sheetsResult.data);
  let sheets = allSheets;
  if (filterSheetName) {
    const hit = allSheets.find(
      (s) => s.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) === 0,
    );
    if (!hit) {
      return fail(
        "workbook.objects.inspect",
        "wps-jsa",
        `Worksheet not found: ${filterSheetName}`,
        "WPS Worksheets.Item by name",
      );
    }
    sheets = [hit];
  }

  const activeName =
    workbook.ActiveSheet?.Name != null ? String(workbook.ActiveSheet.Name) : sheets.find((s) => s.isActive)?.name ?? "";

  // Named ranges: workbook scope + worksheet scopes when members exist.
  const namedItems: NamedRangeInfo[] = [];
  const namedLimitations: string[] = [];
  let namedFailed: { reason: string; evidence?: string } | null = null;

  const wbNames = await deps.listNamedRanges({ scope: "workbook" });
  if (wbNames.ok) {
    namedItems.push(...wbNames.data);
  } else if (wbNames.unsupported) {
    namedLimitations.push(`workbook-scoped Names unavailable: ${wbNames.reason}`);
  } else {
    namedFailed = { reason: wbNames.reason, evidence: wbNames.evidence };
  }

  if (!namedFailed) {
    const sheetTargets = filterSheetName ? sheets : allSheets;
    for (const sheet of sheetTargets) {
      const wsNames = await deps.listNamedRanges({
        scope: "worksheet",
        sheetName: sheet.name,
      });
      if (wsNames.ok) {
        namedItems.push(
          ...wsNames.data.map((n) => ({
            ...n,
            scope: "worksheet" as const,
            sheetName: n.sheetName ?? sheet.name,
          })),
        );
      } else if (wsNames.unsupported) {
        namedLimitations.push(
          `worksheet-scoped Names unavailable on ${sheet.name}: ${wsNames.reason}`,
        );
      } else {
        namedLimitations.push(
          `worksheet-scoped Names failed on ${sheet.name}: ${sanitizeInventoryMessage(wsNames.reason)}`,
        );
      }
    }
  }

  const namedRanges =
    namedFailed != null
      ? failedCategory<NamedRangeInfo>(namedFailed.reason, namedFailed.evidence, namedLimitations)
      : availableCategory(namedItems, maxItems, sortNamedRanges, {
          limitations: namedLimitations.length > 0 ? namedLimitations : undefined,
        });

  const tables = unsupportedCategory<TableInfo>(
    "Table list is not verified for WPS JSA",
    TABLE_EVIDENCE,
  );
  const charts = unsupportedCategory<ChartInfo>(
    "Chart list is not verified for WPS JSA",
    CHART_EVIDENCE,
  );
  const shapes = unsupportedCategory<ShapeInfo>(
    "shape.list is not verified for WPS JSA",
    SHAPE_EVIDENCE,
  );

  if (namedRanges.truncated || namedLimitations.length > 0) {
    limitations.push(...namedLimitations);
  }
  if (namedRanges.truncated) {
    limitations.push(
      `namedRanges truncated at maxItemsPerCategory=${maxItems}; totalCount is the true size.`,
    );
  }

  return ok({
    workbookName: workbook.Name ?? "",
    activeSheetName: activeName,
    sheetCount: allSheets.length,
    sheets,
    tables,
    charts,
    namedRanges,
    shapes,
    limitations,
    filterSheetName,
  });
}
