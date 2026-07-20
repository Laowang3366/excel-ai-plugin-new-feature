/**
 * workbook.objects.inspect — Office.js inventory.
 *
 * Isolation: fixed number of Excel.run calls (1 baseline + 1 per category = 5),
 * independent of sheet/object count. Category sync failures do not reuse a
 * poisoned RequestContext for other categories.
 */
import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import {
  availableCategory,
  buildSheetOrder,
  failedCategory,
  unsupportedCategory,
  sortCharts,
  sortNamedRanges,
  sortShapes,
  sortSheets,
  sortTables,
  type SheetOrderIndex,
} from "./workbookObjectsHelpers";
import type {
  ObjectCategoryResult,
  WorkbookObjectsInspectInfo,
  WorkbookObjectsInspectInput,
} from "./workbookObjectsTypes";
import { WORKBOOK_OBJECTS_MAX_DEFAULT } from "./workbookObjectsTypes";

type SheetProxy = {
  name: string;
  position: number;
  tables: {
    items: Array<{
      name: string;
      showHeaders: boolean;
      showFilterButton: boolean;
      getRange(): { address?: string; load(props: string): void };
    }>;
    load(props: string): void;
  };
  charts: {
    items: Array<{
      name: string;
      chartType: string;
      style?: number;
      left: number;
      top: number;
      width: number;
      height: number;
      title: { text?: string; load(props: string): void };
      legend: { visible?: boolean; load(props: string): void };
    }>;
    load(props: string): void;
  };
  shapes: {
    items: Array<{
      name: string;
      type?: string;
      geometricShapeType?: string | null;
      left: number;
      top: number;
      width: number;
      height: number;
      visible?: boolean;
      load(props: string): void;
    }>;
    load(props: string): void;
  };
  names: {
    items: Array<{ name: string; formula?: string; visible?: boolean }>;
    load(props: string): void;
  };
  load(props: string): void;
};

type ContextLike = {
  workbook: {
    name: string;
    load(props: string): void;
    worksheets: {
      items: SheetProxy[];
      load(props: string): void;
      getActiveWorksheet(): { name: string; load(props: string): void };
      getItem(name: string): SheetProxy;
    };
    names: {
      items: Array<{ name: string; formula?: string; visible?: boolean }>;
      load(props: string): void;
    };
  };
  sync(): Promise<void>;
};

type Baseline = {
  workbookName: string;
  activeSheetName: string;
  allSheets: SheetInfo[];
  targetSheetNames: string[];
  filterSheetName?: string;
};

function resolveTargets(
  allSheets: SheetInfo[],
  filterSheetName: string | undefined,
): { sheets: SheetInfo[]; targetSheetNames: string[] } {
  if (!filterSheetName) {
    return { sheets: allSheets, targetSheetNames: allSheets.map((s) => s.name) };
  }
  const hit = allSheets.find(
    (s) => s.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) === 0,
  );
  if (!hit) throw new Error(`Worksheet not found: ${filterSheetName}`);
  return { sheets: [hit], targetSheetNames: [hit.name] };
}

export async function officeJsInspectWorkbookObjects(
  input: WorkbookObjectsInspectInput = {},
): Promise<HostResult<WorkbookObjectsInspectInfo>> {
  const maxItems = input.maxItemsPerCategory ?? WORKBOOK_OBJECTS_MAX_DEFAULT;
  const filterSheetName = input.sheetName?.trim() || undefined;

  const baselineResult = await withExcel("workbook.objects.inspect", async (context) => {
    const ctx = context as unknown as ContextLike;
    ctx.workbook.load("name");
    ctx.workbook.worksheets.load("items/name,items/position");
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    await ctx.sync();
    const allSheets = sortSheets(
      ctx.workbook.worksheets.items.map((sheet) => ({
        name: sheet.name,
        index: sheet.position,
        isActive: sheet.name === active.name,
      })),
    );
    const resolved = resolveTargets(allSheets, filterSheetName);
    return {
      workbookName: ctx.workbook.name,
      activeSheetName: active.name,
      allSheets,
      targetSheetNames: resolved.targetSheetNames,
      filterSheetName,
    } satisfies Baseline;
  });
  if (!baselineResult.ok) return baselineResult;
  const baseline = baselineResult.data;
  const order = buildSheetOrder(baseline.allSheets);
  const sheets = filterSheetName
    ? baseline.allSheets.filter((s) =>
        s.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) === 0,
      )
    : baseline.allSheets;

  const limitations: string[] = [
    "Inventory is a capped snapshot for model context; use dedicated list tools for full detail.",
    "Shape text is omitted in inventory (geometry/type only) to bound payload size.",
    "Office.js inventory uses one Excel.run for baseline plus one isolated run per object category (fixed bound).",
  ];

  const tables = await runTables(baseline.targetSheetNames, maxItems, order);
  const charts = await runCharts(baseline.targetSheetNames, maxItems, order);
  const namedRanges = await runNamedRanges(
    baseline.allSheets.map((s) => s.name),
    filterSheetName,
    maxItems,
    order,
  );
  const shapes = await runShapes(baseline.targetSheetNames, maxItems, order);

  for (const cat of [tables, charts, namedRanges, shapes]) {
    if (cat.truncated) {
      limitations.push(
        `At least one category truncated at maxItemsPerCategory=${maxItems}; totalCount is the true size.`,
      );
      break;
    }
  }

  return {
    ok: true,
    data: {
      workbookName: baseline.workbookName,
      activeSheetName: baseline.activeSheetName,
      sheetCount: baseline.allSheets.length,
      sheets,
      tables,
      charts,
      namedRanges,
      shapes,
      limitations,
      filterSheetName,
    },
  };
}

async function runTables(
  sheetNames: string[],
  maxItems: number,
  order: SheetOrderIndex,
): Promise<ObjectCategoryResult<TableInfo>> {
  const result = await withExcel("workbook.objects.inspect.tables", async (context) => {
    const ctx = context as unknown as ContextLike;
    const sheets = sheetNames.map((name) => ctx.workbook.worksheets.getItem(name));
    for (const sheet of sheets) {
      sheet.load("name");
      sheet.tables.load("items/name,items/showHeaders,items/showFilterButton");
    }
    await ctx.sync();
    const pending: Array<{
      sheetName: string;
      table: SheetProxy["tables"]["items"][number];
      range: { address?: string; load(props: string): void };
    }> = [];
    for (const sheet of sheets) {
      for (const table of sheet.tables.items) {
        const range = table.getRange();
        range.load("address");
        pending.push({ sheetName: sheet.name, table, range });
      }
    }
    await ctx.sync();
    const items: TableInfo[] = pending.map(({ sheetName, table, range }) => ({
      name: table.name,
      sheetName,
      address: String(range.address ?? ""),
      hasHeaders: table.showHeaders,
      showFilter: table.showFilterButton,
    }));
    return items;
  });
  if (result.ok) {
    return availableCategory(result.data, maxItems, (items) => sortTables(items, order));
  }
  if (result.unsupported) {
    return unsupportedCategory(result.reason, result.evidence);
  }
  return failedCategory(result.reason, "Office.js tables inventory (isolated Excel.run)");
}

async function runCharts(
  sheetNames: string[],
  maxItems: number,
  order: SheetOrderIndex,
): Promise<ObjectCategoryResult<ChartInfo>> {
  const result = await withExcel("workbook.objects.inspect.charts", async (context) => {
    const ctx = context as unknown as ContextLike;
    const sheets = sheetNames.map((name) => ctx.workbook.worksheets.getItem(name));
    for (const sheet of sheets) {
      sheet.load("name");
      sheet.charts.load(
        "items/name,items/chartType,items/style,items/left,items/top,items/width,items/height",
      );
    }
    await ctx.sync();
    for (const sheet of sheets) {
      for (const chart of sheet.charts.items) {
        chart.title.load("text");
        chart.legend.load("visible");
      }
    }
    await ctx.sync();
    const items: ChartInfo[] = [];
    for (const sheet of sheets) {
      for (const chart of sheet.charts.items) {
        items.push({
          name: chart.name,
          sheetName: sheet.name,
          chartType: toChartTypeLabel(String(chart.chartType)),
          title: chart.title.text ?? "",
          left: chart.left,
          top: chart.top,
          width: chart.width,
          height: chart.height,
          style: chart.style,
          legendVisible: chart.legend.visible,
        });
      }
    }
    return items;
  });
  if (result.ok) {
    return availableCategory(result.data, maxItems, (items) => sortCharts(items, order));
  }
  if (result.unsupported) {
    return unsupportedCategory(result.reason, result.evidence);
  }
  return failedCategory(result.reason, "Office.js charts inventory (isolated Excel.run)");
}

async function runNamedRanges(
  allSheetNames: string[],
  filterSheetName: string | undefined,
  maxItems: number,
  order: SheetOrderIndex,
): Promise<ObjectCategoryResult<NamedRangeInfo>> {
  const result = await withExcel("workbook.objects.inspect.namedRanges", async (context) => {
    const ctx = context as unknown as ContextLike;
    ctx.workbook.names.load("items/name,items/formula,items/visible");
    const sheets = allSheetNames.map((name) => ctx.workbook.worksheets.getItem(name));
    for (const sheet of sheets) {
      sheet.load("name");
      sheet.names.load("items/name,items/formula,items/visible");
    }
    await ctx.sync();
    const items: NamedRangeInfo[] = [];
    for (const item of ctx.workbook.names.items) {
      items.push({
        name: item.name,
        refersTo: item.formula ?? "",
        scope: "workbook",
        visible: item.visible,
      });
    }
    for (const sheet of sheets) {
      if (
        filterSheetName &&
        sheet.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) !== 0
      ) {
        continue;
      }
      for (const item of sheet.names.items) {
        items.push({
          name: item.name,
          refersTo: item.formula ?? "",
          scope: "worksheet",
          sheetName: sheet.name,
          visible: item.visible,
        });
      }
    }
    return items;
  });
  if (result.ok) {
    return availableCategory(result.data, maxItems, (items) => sortNamedRanges(items, order));
  }
  if (result.unsupported) {
    return unsupportedCategory(result.reason, result.evidence);
  }
  return failedCategory(result.reason, "Office.js named ranges inventory (isolated Excel.run)");
}

async function runShapes(
  sheetNames: string[],
  maxItems: number,
  order: SheetOrderIndex,
): Promise<ObjectCategoryResult<ShapeInfo>> {
  const result = await withExcel("workbook.objects.inspect.shapes", async (context) => {
    const ctx = context as unknown as ContextLike;
    const sheets = sheetNames.map((name) => ctx.workbook.worksheets.getItem(name));
    for (const sheet of sheets) {
      sheet.load("name");
      sheet.shapes.load(
        "items/name,items/type,items/geometricShapeType,items/left,items/top,items/width,items/height,items/visible",
      );
    }
    await ctx.sync();
    const items: ShapeInfo[] = [];
    for (const sheet of sheets) {
      for (const shape of sheet.shapes.items) {
        items.push({
          name: shape.name,
          sheetName: sheet.name,
          type: String(shape.type ?? "Unsupported"),
          geometricShapeType:
            shape.geometricShapeType == null ? null : String(shape.geometricShapeType),
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
          visible: shape.visible,
          text: null,
        });
      }
    }
    return items;
  });
  if (result.ok) {
    return availableCategory(result.data, maxItems, (items) => sortShapes(items, order), {
      limitations: ["Shape text omitted in workbook.objects.inspect inventory"],
    });
  }
  if (result.unsupported) {
    return unsupportedCategory(result.reason, result.evidence);
  }
  return failedCategory(result.reason, "Office.js shapes inventory (isolated Excel.run)");
}
