/** @MOCK_INTERFACE — Excel.run table/chart test double with write→sync→load→sync fence. */

type TableState = {
  name: string;
  sheetName: string;
  showHeaders: boolean;
  showFilterButton: boolean;
  showTotals: boolean;
  showBandedRows: boolean;
  showBandedColumns: boolean;
  style: string;
  address: string;
};

type TableEntry = { committed: TableState; pending: Partial<TableState> | undefined };

type ChartState = {
  name: string;
  chartType: string;
  title: string;
  titleVisible: boolean;
  style: number;
  legendVisible: boolean;
  left: number;
  top: number;
  width: number;
  height: number;
};

type ChartEntry = { committed: ChartState; pending: Partial<ChartState> | undefined };

export function installObjectUpdateExcel() {
  const tables = new Map<string, TableEntry>();
  tables.set("T1", {
    committed: {
      name: "T1",
      sheetName: "Sheet1",
      showHeaders: true,
      showFilterButton: true,
      showTotals: false,
      showBandedRows: true,
      showBandedColumns: false,
      style: "TableStyleMedium2",
      address: "Sheet1!A1:C3",
    },
    pending: undefined,
  });
  const charts = new Map<string, ChartEntry>();
  charts.set("C1", {
    committed: {
      name: "C1",
      chartType: "ColumnClustered",
      title: "Old",
      titleVisible: true,
      style: 2,
      legendVisible: true,
      left: 10,
      top: 20,
      width: 300,
      height: 200,
    },
    pending: undefined,
  });
  function makeTable(name: string) {
    const entry = tables.get(name);
    if (!entry) throw new Error(`missing table ${name}`);
    let pendingSnap: TableState | undefined;
    let snapshot: TableState | undefined;
    let pendingAddress: string | undefined;
    let address: string | undefined;
    const proxy = {
      get name() {
        if (!snapshot) throw new Error("Table.name not loaded");
        return snapshot.name;
      },
      set name(next: string) {
        entry.pending = { ...entry.pending, name: next };
      },
      get showHeaders() {
        if (!snapshot) throw new Error("Table.showHeaders not loaded");
        return snapshot.showHeaders;
      },
      set showHeaders(v: boolean) {
        entry.pending = { ...entry.pending, showHeaders: v };
      },
      get showFilterButton() {
        if (!snapshot) throw new Error("Table.showFilterButton not loaded");
        return snapshot.showFilterButton;
      },
      set showFilterButton(v: boolean) {
        entry.pending = { ...entry.pending, showFilterButton: v };
      },
      get showTotals() {
        if (!snapshot) throw new Error("Table.showTotals not loaded");
        return snapshot.showTotals;
      },
      set showTotals(v: boolean) {
        entry.pending = { ...entry.pending, showTotals: v };
      },
      get showBandedRows() {
        if (!snapshot) throw new Error("Table.showBandedRows not loaded");
        return snapshot.showBandedRows;
      },
      set showBandedRows(v: boolean) {
        entry.pending = { ...entry.pending, showBandedRows: v };
      },
      get showBandedColumns() {
        if (!snapshot) throw new Error("Table.showBandedColumns not loaded");
        return snapshot.showBandedColumns;
      },
      set showBandedColumns(v: boolean) {
        entry.pending = { ...entry.pending, showBandedColumns: v };
      },
      get style() {
        if (!snapshot) throw new Error("Table.style not loaded");
        return snapshot.style;
      },
      set style(v: string) {
        entry.pending = { ...entry.pending, style: v };
      },
      resize(next: string) {
        const currentRow = /![A-Z]+(\d+)/i.exec(entry.committed.address)?.[1];
        const nextRow = /^[A-Z]+(\d+)/i.exec(next)?.[1];
        if (nextRow !== currentRow) {
          throw new Error("Table.resize requires the header row to remain on the same row");
        }
        entry.pending = {
          ...entry.pending,
          address: next.includes("!") ? next : `${entry.committed.sheetName}!${next}`,
        };
      },
      getRange() {
        return {
          get address() {
            if (!address) throw new Error("Table range address not loaded");
            return address;
          },
          load() {
            pendingAddress = entry.committed.address;
            address = undefined;
          },
        };
      },
      load() {
        pendingSnap = { ...entry.committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
        if (pendingAddress !== undefined) {
          address = pendingAddress;
          pendingAddress = undefined;
        }
      },
      _entry: entry,
    };
    return proxy;
  }

  const tableProxyByName = new Map<string, ReturnType<typeof makeTable>>();

  function tableFor(name: string) {
    let proxy = tableProxyByName.get(name);
    if (!proxy) {
      proxy = makeTable(name);
      tableProxyByName.set(name, proxy);
    }
    return proxy;
  }

  function makeChart(name: string) {
    const entry = charts.get(name);
    if (!entry) throw new Error(`missing chart ${name}`);

    let pendingSnap: ChartState | undefined;
    let snapshot: ChartState | undefined;
    let pendingTitleSnap: { text: string; visible: boolean } | undefined;
    let titleSnap: { text: string; visible: boolean } | undefined;
    let pendingLegendSnap: boolean | undefined;
    let legendSnap: boolean | undefined;

    const proxy = {
      get name() {
        if (!snapshot) throw new Error("Chart.name not loaded");
        return snapshot.name;
      },
      set name(next: string) {
        entry.pending = { ...entry.pending, name: next };
      },
      get chartType() {
        if (!snapshot) throw new Error("Chart.chartType not loaded");
        return snapshot.chartType;
      },
      set chartType(v: string) {
        entry.pending = { ...entry.pending, chartType: v };
      },
      get style() {
        if (!snapshot) throw new Error("Chart.style not loaded");
        return snapshot.style;
      },
      set style(v: number) {
        entry.pending = { ...entry.pending, style: v };
      },
      title: {
        get text() {
          if (!titleSnap) throw new Error("Chart.title.text not loaded");
          return titleSnap.text;
        },
        set text(v: string) {
          entry.pending = { ...entry.pending, title: v };
        },
        get visible() {
          if (!titleSnap) throw new Error("Chart.title.visible not loaded");
          return titleSnap.visible;
        },
        set visible(v: boolean) {
          entry.pending = { ...entry.pending, titleVisible: v };
        },
        load() {
          pendingTitleSnap = {
            text: entry.committed.title,
            visible: entry.committed.titleVisible,
          };
          titleSnap = undefined;
        },
      },
      legend: {
        get visible() {
          if (legendSnap === undefined) throw new Error("Chart.legend.visible not loaded");
          return legendSnap;
        },
        set visible(v: boolean) {
          entry.pending = { ...entry.pending, legendVisible: v };
        },
        load() {
          pendingLegendSnap = entry.committed.legendVisible;
          legendSnap = undefined;
        },
      },
      get left() {
        if (!snapshot) throw new Error("Chart.left not loaded");
        return snapshot.left;
      },
      set left(v: number) {
        entry.pending = { ...entry.pending, left: v };
      },
      get top() {
        if (!snapshot) throw new Error("Chart.top not loaded");
        return snapshot.top;
      },
      set top(v: number) {
        entry.pending = { ...entry.pending, top: v };
      },
      get width() {
        if (!snapshot) throw new Error("Chart.width not loaded");
        return snapshot.width;
      },
      set width(v: number) {
        entry.pending = { ...entry.pending, width: v };
      },
      get height() {
        if (!snapshot) throw new Error("Chart.height not loaded");
        return snapshot.height;
      },
      set height(v: number) {
        entry.pending = { ...entry.pending, height: v };
      },
      load() {
        pendingSnap = { ...entry.committed };
        snapshot = undefined;
      },
      _flushLoad() {
        if (pendingSnap) {
          snapshot = pendingSnap;
          pendingSnap = undefined;
        }
        if (pendingTitleSnap) {
          titleSnap = pendingTitleSnap;
          pendingTitleSnap = undefined;
        }
        if (pendingLegendSnap !== undefined) {
          legendSnap = pendingLegendSnap;
          pendingLegendSnap = undefined;
        }
      },
      _entry: entry,
    };
    return proxy;
  }

  const chartProxyByName = new Map<string, ReturnType<typeof makeChart>>();

  function chartFor(name: string) {
    let proxy = chartProxyByName.get(name);
    if (!proxy) {
      proxy = makeChart(name);
      chartProxyByName.set(name, proxy);
    }
    return proxy;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(_name: string) {
          return {
            tables: {
              getItem(tableName: string) {
                return tableFor(tableName);
              },
            },
            charts: {
              getItem(chartName: string) {
                return chartFor(chartName);
              },
            },
          };
        },
      },
    },
    async sync() {
      for (const entry of tables.values()) {
        if (!entry.pending) continue;
        const oldName = entry.committed.name;
        Object.assign(entry.committed, entry.pending);
        entry.pending = undefined;
        if (entry.committed.name !== oldName) {
          tables.delete(oldName);
          tables.set(entry.committed.name, entry);
          const proxy = tableProxyByName.get(oldName);
          if (proxy) {
            tableProxyByName.delete(oldName);
            tableProxyByName.set(entry.committed.name, proxy);
          }
        }
      }
      for (const entry of charts.values()) {
        if (entry.pending) {
          const oldName = entry.committed.name;
          Object.assign(entry.committed, entry.pending);
          entry.pending = undefined;
          if (entry.committed.name !== oldName) {
            charts.delete(oldName);
            charts.set(entry.committed.name, entry);
            const proxy = chartProxyByName.get(oldName);
            if (proxy) {
              chartProxyByName.delete(oldName);
              chartProxyByName.set(entry.committed.name, proxy);
            }
          }
        }
      }
      for (const proxy of tableProxyByName.values()) proxy._flushLoad();
      for (const proxy of chartProxyByName.values()) {
        proxy._flushLoad();
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as {
    Excel: { run: Function; ChartType: Record<string, string> };
  }).Excel = {
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
    tables,
    charts,
    /** write → load → single sync yields stale chartType (skip first sync). */
    async brokenUpdateSkipFirstSync(chartType: string) {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const chart = sheet.charts.getItem("C1");
      chart.chartType = chartType;
      chart.load();
      chart.title.load();
      chart.legend.load();
      await context.sync();
      return { chartType: chart.chartType, title: chart.title.text };
    },
  };
}
