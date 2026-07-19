import type { CellValue } from "../../shared/host/types";
import {
  createSyncGatedChartHost,
  type ChartState,
} from "./officeJsSyncGatedChart";

type CellState = {
  values: CellValue[][];
  formulas: string[][];
  format: Record<string, unknown>;
  numberFormat: string;
};

type SheetState = {
  name: string;
  position: number;
  cells: Map<string, CellState>;
  tables: Map<string, { name: string; address: string; showHeaders: boolean }>;
  charts: Map<string, ChartState>;
};

/**
 * Fake Excel.run where collection.items stay empty until load()+sync().
 * Proves listCharts must not capture items before sync.
 * Chart create/list readback state machine lives in officeJsSyncGatedChart.
 */
function installSyncGatedExcel() {
  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", {
    name: "Sheet1",
    position: 0,
    cells: new Map(),
    tables: new Map(),
    charts: new Map(),
  });

  let worksheetItems: ReturnType<typeof makeSheetApi>[] = [];
  let pendingWorksheetItems: ReturnType<typeof makeSheetApi>[] | null = null;

  function getSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing ${name}`);
    return sheet;
  }

  const chartHost = createSyncGatedChartHost({
    getSheet,
    sheets,
  });

  function makeRange(sheetName: string, address: string) {
    const sheet = getSheet(sheetName);
    const key = address.toUpperCase();
    if (!sheet.cells.has(key)) {
      sheet.cells.set(key, {
        values: [[null]],
        formulas: [[""]],
        format: {
          font: { name: "Calibri", size: 11, bold: false, color: "#000000" },
          fill: { color: "#FFFFFF" },
          horizontalAlignment: "General",
          verticalAlignment: "Bottom",
          wrapText: false,
        },
        numberFormat: "General",
      });
    }
    const state = () => sheet.cells.get(key)!;
    const range: Record<string, unknown> = {
      address: `${sheetName}!${address}`,
      rowCount: 2,
      columnCount: 2,
      get values() {
        return state().values;
      },
      set values(next: CellValue[][]) {
        state().values = next;
      },
      get formulas() {
        return state().formulas;
      },
      set formulas(next: string[][]) {
        state().formulas = next;
      },
      get numberFormat() {
        return [[state().numberFormat]];
      },
      set numberFormat(next: string[][]) {
        state().numberFormat = next[0]?.[0] ?? "General";
      },
      format: {
        get font() {
          return state().format.font as Record<string, unknown>;
        },
        get fill() {
          return state().format.fill as Record<string, unknown>;
        },
        get horizontalAlignment() {
          return (state().format as { horizontalAlignment: string }).horizontalAlignment;
        },
        set horizontalAlignment(v: string) {
          (state().format as { horizontalAlignment: string }).horizontalAlignment = v;
        },
        get verticalAlignment() {
          return (state().format as { verticalAlignment: string }).verticalAlignment;
        },
        set verticalAlignment(v: string) {
          (state().format as { verticalAlignment: string }).verticalAlignment = v;
        },
        get wrapText() {
          return (state().format as { wrapText: boolean }).wrapText;
        },
        set wrapText(v: boolean) {
          (state().format as { wrapText: boolean }).wrapText = v;
        },
        load() {},
      },
      load() {},
      clear() {
        sheet.cells.delete(key);
      },
    };
    (range.format as { font: { load: () => void }; fill: { load: () => void } }).font.load =
      () => {};
    (range.format as { fill: { load: () => void } }).fill.load = () => {};
    return range;
  }

  function makeSheetApi(name: string) {
    const sheet = getSheet(name);
    return {
      get name() {
        return sheet.name;
      },
      set name(next: string) {
        sheets.delete(sheet.name);
        sheet.name = next;
        sheets.set(next, sheet);
      },
      position: sheet.position,
      load() {},
      getRange(address: string) {
        return makeRange(sheet.name, address);
      },
      getUsedRangeOrNullObject() {
        const range = makeRange(sheet.name, "A1:B2") as {
          address: string;
          isNullObject?: boolean;
          load: () => void;
        };
        range.isNullObject = false;
        return range;
      },
      delete() {
        sheets.delete(sheet.name);
      },
      tables: {
        get items() {
          return [...sheet.tables.values()].map((table) => ({
            name: table.name,
            showHeaders: table.showHeaders,
            showFilterButton: true,
            load() {},
            getRange() {
              return makeRange(sheet.name, table.address);
            },
            delete() {
              sheet.tables.delete(table.name);
            },
          }));
        },
        load() {},
        add(address: string, hasHeaders: boolean) {
          const tableName = `Table${sheet.tables.size + 1}`;
          const table = { name: tableName, address, showHeaders: hasHeaders };
          sheet.tables.set(tableName, table);
          return {
            get name() {
              return table.name;
            },
            set name(next: string) {
              sheet.tables.delete(table.name);
              table.name = next;
              sheet.tables.set(next, table);
            },
            showHeaders: hasHeaders,
            showFilterButton: true,
            load() {},
            getRange() {
              return makeRange(sheet.name, address);
            },
            delete() {
              sheet.tables.delete(table.name);
            },
          };
        },
        getItem(tableName: string) {
          const table = sheet.tables.get(tableName);
          if (!table) throw new Error("missing table");
          return {
            name: table.name,
            showHeaders: table.showHeaders,
            showFilterButton: true,
            load() {},
            getRange() {
              return makeRange(sheet.name, table.address);
            },
            delete() {
              sheet.tables.delete(table.name);
            },
          };
        },
      },
      charts: chartHost.makeChartsApi(sheet.name),
    };
  }

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load() {},
      worksheets: {
        get items() {
          return worksheetItems;
        },
        load(_props?: string) {
          pendingWorksheetItems = [...sheets.keys()].map((name) => makeSheetApi(name));
        },
        getActiveWorksheet() {
          return makeSheetApi([...sheets.keys()][0]!);
        },
        getItem(name: string) {
          return makeSheetApi(name);
        },
        add(name?: string) {
          const sheetName = name ?? `Sheet${sheets.size + 1}`;
          sheets.set(sheetName, {
            name: sheetName,
            position: sheets.size,
            cells: new Map(),
            tables: new Map(),
            charts: new Map(),
          });
          return makeSheetApi(sheetName);
        },
      },
      getSelectedRange() {
        return makeRange("Sheet1", "A1");
      },
    },
    async sync() {
      chartHost.commitPendingWrites();
      if (pendingWorksheetItems) {
        worksheetItems = pendingWorksheetItems;
        pendingWorksheetItems = null;
      }
      chartHost.applyChartSnapshots();
      chartHost.flushLoads();
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function; ChartType: Record<string, string> } }).Excel =
    {
      run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
      ChartType: {
        columnClustered: "ColumnClustered",
        line: "Line",
        barClustered: "BarClustered",
        area: "Area",
        pie: "Pie",
        xyscatter: "XYScatter",
        doughnut: "Doughnut",
        bubble: "Bubble",
        radar: "Radar",
        lineMarkers: "LineMarkers",
      },
    };

  return {
    /** Capture items the broken way (before sync) for regression evidence. */
    captureItemsBeforeSync() {
      context.workbook.worksheets.load("items/name");
      return context.workbook.worksheets.items.slice();
    },
    async captureItemsAfterSync() {
      context.workbook.worksheets.load("items/name");
      await context.sync();
      return context.workbook.worksheets.items.slice();
    },
    getChartTitleVisible(sheetName: string, chartName: string): boolean | undefined {
      return sheets.get(sheetName)?.charts.get(chartName)?.titleVisible;
    },
    /** write → load → single sync (skip first sync) yields stale chartType. */
    async brokenCreateSkipFirstSync(chartType: string) {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const chart = sheet.charts.add(chartType, sheet.getRange("A1:B2"));
      chart.load();
      chart.title.load();
      chart.legend.load();
      await context.sync();
      return {
        name: chart.name,
        chartType: chart.chartType,
      };
    },
  };
}

export { installSyncGatedExcel };
