/**
 * Batched fake for workbook.objects.inspect (multi sheet / table / chart / name / shape).
 */
export function installWorkbookObjectsExcel(options?: {
  failTables?: boolean;
  failCharts?: boolean;
  failNames?: boolean;
  failShapes?: boolean;
  failWorkbook?: boolean;
}) {
  type Named = { name: string; formula: string; visible: boolean };
  type Table = { name: string; showHeaders: boolean; showFilterButton: boolean; address: string };
  type Chart = {
    name: string;
    chartType: string;
    style: number;
    left: number;
    top: number;
    width: number;
    height: number;
    title: string;
    legendVisible: boolean;
  };
  type Shape = {
    name: string;
    type: string;
    geometricShapeType: string | null;
    left: number;
    top: number;
    width: number;
    height: number;
    visible: boolean;
  };
  type Sheet = {
    name: string;
    position: number;
    tables: Table[];
    charts: Chart[];
    shapes: Shape[];
    names: Named[];
  };

  const sheets = new Map<string, Sheet>();
  sheets.set("Sheet1", {
    name: "Sheet1",
    position: 0,
    tables: [
      { name: "T_B", showHeaders: true, showFilterButton: true, address: "Sheet1!A1:B2" },
      { name: "T_A", showHeaders: true, showFilterButton: false, address: "Sheet1!C1:D2" },
    ],
    charts: [
      {
        name: "ChartB",
        chartType: "Line",
        style: 10,
        left: 1,
        top: 2,
        width: 100,
        height: 80,
        title: "B",
        legendVisible: true,
      },
      {
        name: "ChartA",
        chartType: "ColumnClustered",
        style: 8,
        left: 0,
        top: 0,
        width: 120,
        height: 90,
        title: "A",
        legendVisible: false,
      },
    ],
    shapes: [
      {
        name: "ShapeB",
        type: "GeometricShape",
        geometricShapeType: "Ellipse",
        left: 0,
        top: 0,
        width: 40,
        height: 40,
        visible: true,
      },
      {
        name: "ShapeA",
        type: "GeometricShape",
        geometricShapeType: "Rectangle",
        left: 10,
        top: 10,
        width: 50,
        height: 50,
        visible: true,
      },
    ],
    names: [{ name: "LocalZ", formula: "=Sheet1!$A$1", visible: true }],
  });
  sheets.set("Data", {
    name: "Data",
    position: 1,
    tables: [{ name: "T_Data", showHeaders: true, showFilterButton: true, address: "Data!A1:C3" }],
    charts: [],
    shapes: [
      {
        name: "DataShape",
        type: "GeometricShape",
        geometricShapeType: "Triangle",
        left: 5,
        top: 5,
        width: 30,
        height: 30,
        visible: true,
      },
    ],
    names: [{ name: "DataLocal", formula: "=Data!$B$2", visible: true }],
  });

  const workbookNames: Named[] = [
    { name: "WbName", formula: "=Sheet1!$Z$1", visible: true },
    { name: "Alpha", formula: "=1", visible: false },
  ];

  function makeNamedCollection(list: Named[]) {
    return {
      items: list.map((n) => ({
        name: n.name,
        formula: n.formula,
        visible: n.visible,
      })),
      load() {
        if (options?.failNames) throw new Error("names load failed");
      },
    };
  }

  function makeSheetProxy(sheet: Sheet) {
    return {
      get name() {
        return sheet.name;
      },
      get position() {
        return sheet.position;
      },
      load() {},
      tables: {
        get items() {
          return sheet.tables.map((t) => ({
            name: t.name,
            showHeaders: t.showHeaders,
            showFilterButton: t.showFilterButton,
            getRange() {
              return {
                address: t.address,
                load() {},
              };
            },
          }));
        },
        load() {
          if (options?.failTables) throw new Error("tables load failed");
        },
      },
      charts: {
        get items() {
          return sheet.charts.map((c) => ({
            name: c.name,
            chartType: c.chartType,
            style: c.style,
            left: c.left,
            top: c.top,
            width: c.width,
            height: c.height,
            title: {
              text: c.title,
              load() {},
            },
            legend: {
              visible: c.legendVisible,
              load() {},
            },
          }));
        },
        load() {
          if (options?.failCharts) throw new Error("charts load failed");
        },
      },
      shapes: {
        get items() {
          return sheet.shapes.map((s) => ({
            name: s.name,
            type: s.type,
            geometricShapeType: s.geometricShapeType,
            left: s.left,
            top: s.top,
            width: s.width,
            height: s.height,
            visible: s.visible,
            load() {},
          }));
        },
        load() {
          if (options?.failShapes) throw new Error("shapes load failed");
        },
      },
      names: makeNamedCollection(sheet.names),
    };
  }

  const context = {
    workbook: {
      name: "Book1.xlsx",
      load() {
        if (options?.failWorkbook) throw new Error("workbook load failed");
      },
      names: makeNamedCollection(workbookNames),
      worksheets: {
        get items() {
          return [...sheets.values()].map(makeSheetProxy);
        },
        load() {},
        getActiveWorksheet() {
          return { name: "Sheet1", load() {} };
        },
        getItem(name: string) {
          const sheet = sheets.get(name);
          if (!sheet) throw new Error(`ItemNotFound: Worksheet ${name} not found`);
          return makeSheetProxy(sheet);
        },
      },
    },
    async sync() {},
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    context,
  };
}
