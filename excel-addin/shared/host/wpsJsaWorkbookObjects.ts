/**
 * workbook.objects.inspect — WPS JSA partial inventory.
 * Sheets + named ranges via member-probe; table/chart/shape typed unsupported.
 */
import { requireWorkbook } from "./wpsJsaRuntime";
import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import { fail, ok } from "./types";
import {
  availableCategory,
  buildSheetOrder,
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

type NameSource =
  | { kind: "ok"; items: NamedRangeInfo[]; label: string }
  | { kind: "unsupported"; reason: string; evidence?: string; label: string }
  | { kind: "failed"; reason: string; evidence?: string; label: string };

function mapNameResult(
  label: string,
  result: HostResult<NamedRangeInfo[]>,
): NameSource {
  if (result.ok) return { kind: "ok", items: result.data, label };
  if (result.unsupported) {
    return {
      kind: "unsupported",
      reason: result.reason,
      evidence: result.evidence,
      label,
    };
  }
  return { kind: "failed", reason: result.reason, evidence: result.evidence, label };
}

function mergeNamedRanges(
  sources: NameSource[],
  maxItems: number,
  sheetOrder: ReturnType<typeof buildSheetOrder>,
) {
  const okSources = sources.filter((s): s is Extract<NameSource, { kind: "ok" }> => s.kind === "ok");
  const failed = sources.filter((s) => s.kind === "failed");
  const unsupported = sources.filter((s) => s.kind === "unsupported");
  const gapNotes = [...unsupported, ...failed].map(
    (s) => `${s.label}: ${sanitizeInventoryMessage(s.reason)}`,
  );

  if (okSources.length > 0) {
    const items = okSources.flatMap((s) => s.items);
    return availableCategory(items, maxItems, (list) => sortNamedRanges(list, sheetOrder), {
      limitations: gapNotes.length > 0 ? gapNotes : undefined,
    });
  }
  if (failed.length > 0) {
    const first = failed[0]!;
    return failedCategory<NamedRangeInfo>(
      first.reason,
      first.evidence,
      gapNotes,
    );
  }
  if (unsupported.length > 0) {
    const first = unsupported[0]!;
    return unsupportedCategory<NamedRangeInfo>(
      first.reason,
      first.evidence,
      gapNotes,
    );
  }
  return unsupportedCategory<NamedRangeInfo>(
    "No named-range sources were queried",
    "WPS Names inventory",
  );
}

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
    workbook.ActiveSheet?.Name != null
      ? String(workbook.ActiveSheet.Name)
      : sheets.find((s) => s.isActive)?.name ?? "";

  const sources: NameSource[] = [];
  sources.push(
    mapNameResult("workbook.Names", await deps.listNamedRanges({ scope: "workbook" })),
  );
  const sheetTargets = filterSheetName ? sheets : allSheets;
  for (const sheet of sheetTargets) {
    sources.push(
      mapNameResult(
        `worksheet.Names(${sheet.name})`,
        await deps.listNamedRanges({ scope: "worksheet", sheetName: sheet.name }),
      ),
    );
  }

  const sheetOrder = buildSheetOrder(allSheets);
  const namedRanges = mergeNamedRanges(sources, maxItems, sheetOrder);
  if (namedRanges.limitations?.length) {
    limitations.push(...namedRanges.limitations);
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
    tables: unsupportedCategory<TableInfo>(
      "Table list is not verified for WPS JSA",
      TABLE_EVIDENCE,
    ),
    charts: unsupportedCategory<ChartInfo>(
      "Chart list is not verified for WPS JSA",
      CHART_EVIDENCE,
    ),
    namedRanges,
    shapes: unsupportedCategory<ShapeInfo>(
      "shape.list is not verified for WPS JSA",
      SHAPE_EVIDENCE,
    ),
    limitations,
    filterSheetName,
  });
}
