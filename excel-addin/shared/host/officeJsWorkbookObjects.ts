/**
 * workbook.objects.inspect — Office.js batched inventory.
 * One Excel.run; categories load/sync independently so one family failure
 * does not force empty results for others. Avoids per-object unbounded
 * Excel.run (unlike naively calling listTables + listCharts + …).
 */
import { toChartTypeLabel } from "./officeJsChartTypes";
import { withExcel } from "./officeJsRuntime";
import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { HostResult, NamedRangeInfo, SheetInfo, TableInfo } from "./types";
import { fail } from "./types";
import {
  availableCategory,
  failedCategory,
  sanitizeInventoryMessage,
  sortCharts,
  sortNamedRanges,
  sortShapes,
  sortSheets,
  sortTables,
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
      load?(props: string): void;
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

function messageOf(error: unknown): string {
  return sanitizeInventoryMessage(error instanceof Error ? error.message : String(error));
}

export async function officeJsInspectWorkbookObjects(
  input: WorkbookObjectsInspectInput = {},
): Promise<HostResult<WorkbookObjectsInspectInfo>> {
  const maxItems = input.maxItemsPerCategory ?? WORKBOOK_OBJECTS_MAX_DEFAULT;
  const filterSheetName = input.sheetName?.trim() || undefined;

  try {
    return await withExcel("workbook.objects.inspect", async (context) => {
      const ctx = context as unknown as ContextLike;
      const limitations: string[] = [
        "Inventory is a capped snapshot for model context; use dedicated list tools for full detail.",
        "Shape text is omitted in inventory (geometry/type only) to bound payload size.",
      ];

      ctx.workbook.load("name");
      ctx.workbook.worksheets.load("items/name,items/position");
      const active = ctx.workbook.worksheets.getActiveWorksheet();
      active.load("name");
      await ctx.sync();

      const allSheets: SheetInfo[] = sortSheets(
        ctx.workbook.worksheets.items.map((sheet) => ({
          name: sheet.name,
          index: sheet.position,
          isActive: sheet.name === active.name,
        })),
      );

      let sheets = allSheets;
      let targetProxies = [...ctx.workbook.worksheets.items];
      if (filterSheetName) {
        const hit = allSheets.find(
          (s) => s.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) === 0,
        );
        if (!hit) {
          throw new Error(`Worksheet not found: ${filterSheetName}`);
        }
        sheets = [hit];
        targetProxies = ctx.workbook.worksheets.items.filter(
          (s) => s.name.localeCompare(filterSheetName, undefined, { sensitivity: "accent" }) === 0,
        );
      }

      const tables = await collectTables(ctx, targetProxies, maxItems);
      const charts = await collectCharts(ctx, targetProxies, maxItems);
      const namedRanges = await collectNamedRanges(
        ctx,
        filterSheetName ? targetProxies : ctx.workbook.worksheets.items,
        filterSheetName,
        maxItems,
      );
      const shapes = await collectShapes(ctx, targetProxies, maxItems);

      for (const cat of [tables, charts, namedRanges, shapes]) {
        if (cat.truncated) {
          limitations.push(
            `At least one category truncated at maxItemsPerCategory=${maxItems}; totalCount is the true size.`,
          );
          break;
        }
      }

      return {
        workbookName: ctx.workbook.name,
        activeSheetName: active.name,
        sheetCount: allSheets.length,
        sheets,
        tables,
        charts,
        namedRanges,
        shapes,
        limitations,
        filterSheetName,
      };
    });
  } catch (error) {
    return fail(
      "workbook.objects.inspect",
      "office-js",
      messageOf(error),
      "Office.js workbook.objects.inspect base workbook/sheet load",
    );
  }
}

async function collectTables(
  ctx: ContextLike,
  sheets: SheetProxy[],
  maxItems: number,
): Promise<ObjectCategoryResult<TableInfo>> {
  try {
    for (const sheet of sheets) {
      sheet.tables.load("items/name,items/showHeaders,items/showFilterButton");
    }
    await ctx.sync();

    type Pending = {
      sheetName: string;
      table: SheetProxy["tables"]["items"][number];
      range: { address?: string; load(props: string): void };
    };
    const pending: Pending[] = [];
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
    return availableCategory(items, maxItems, sortTables);
  } catch (error) {
    return failedCategory(messageOf(error), "Office.js tables inventory");
  }
}

async function collectCharts(
  ctx: ContextLike,
  sheets: SheetProxy[],
  maxItems: number,
): Promise<ObjectCategoryResult<ChartInfo>> {
  try {
    for (const sheet of sheets) {
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
    return availableCategory(items, maxItems, sortCharts);
  } catch (error) {
    return failedCategory(messageOf(error), "Office.js charts inventory");
  }
}

async function collectNamedRanges(
  ctx: ContextLike,
  sheets: SheetProxy[],
  filterSheetName: string | undefined,
  maxItems: number,
): Promise<ObjectCategoryResult<NamedRangeInfo>> {
  try {
    ctx.workbook.names.load("items/name,items/formula,items/visible");
    for (const sheet of sheets) {
      sheet.names.load("items/name,items/formula,items/visible");
    }
    await ctx.sync();

    const items: NamedRangeInfo[] = [];
    // Workbook scope always included (even with sheet filter).
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
    return availableCategory(items, maxItems, sortNamedRanges);
  } catch (error) {
    return failedCategory(messageOf(error), "Office.js named ranges inventory");
  }
}

async function collectShapes(
  ctx: ContextLike,
  sheets: SheetProxy[],
  maxItems: number,
): Promise<ObjectCategoryResult<ShapeInfo>> {
  try {
    for (const sheet of sheets) {
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
    return availableCategory(items, maxItems, sortShapes, {
      limitations: ["Shape text omitted in workbook.objects.inspect inventory"],
    });
  } catch (error) {
    return failedCategory(messageOf(error), "Office.js shapes inventory");
  }
}
