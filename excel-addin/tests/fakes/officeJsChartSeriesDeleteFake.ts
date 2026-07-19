/**
 * Sync-gated fake for ChartSeries.delete().
 * delete() only marks pending; committed collection changes on first context.sync().
 * load() snapshots committed-at-load; skip-delete-sync (delete→load→sync) keeps old items.
 */

type SeriesState = { name: string; chartType: string; smooth: boolean };

type Entry = {
  committed: SeriesState | null;
  pending: "delete" | undefined;
  order: number;
};

export function installChartSeriesDeleteExcel() {
  const entries: Entry[] = [
    {
      committed: { name: "S1", chartType: "ColumnClustered", smooth: false },
      pending: undefined,
      order: 0,
    },
    {
      committed: { name: "S2", chartType: "ColumnClustered", smooth: false },
      pending: undefined,
      order: 1,
    },
    {
      committed: { name: "S3", chartType: "Line", smooth: true },
      pending: undefined,
      order: 2,
    },
  ];

  let seriesItems: SeriesState[] | null = null;
  let pendingSeriesItems: SeriesState[] | null = null;
  let chartNameValue: unknown = "C1";

  function committedList(): Entry[] {
    return entries
      .filter((e) => e.committed != null)
      .sort((a, b) => a.order - b.order);
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            charts: {
              getItem(chartName: string) {
                if (chartName !== "C1") throw new Error(`missing chart ${chartName}`);
                return {
                  get name() {
                    return chartNameValue as string;
                  },
                  load() {},
                  series: {
                    get items() {
                      return seriesItems ?? [];
                    },
                    getItemAt(index: number) {
                      const list = committedList();
                      if (index < 0 || index >= list.length) {
                        throw new Error(`series index out of range: ${index}`);
                      }
                      const entry = list[index]!;
                      return {
                        delete() {
                          entry.pending = "delete";
                        },
                      };
                    },
                    load(_props?: string) {
                      pendingSeriesItems = committedList().map((e) => ({
                        ...e.committed!,
                      }));
                    },
                  },
                };
              },
            },
          };
        },
      },
    },
    async sync() {
      for (const entry of entries) {
        if (entry.pending === "delete") {
          entry.committed = null;
          entry.pending = undefined;
        }
      }
      if (pendingSeriesItems) {
        seriesItems = pendingSeriesItems;
        pendingSeriesItems = null;
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
      doughnut: "Doughnut",
      bubble: "Bubble",
      radar: "Radar",
      lineMarkers: "LineMarkers",
    },
  };

  return {
    getCommittedNames() {
      return committedList().map((e) => e.committed!.name);
    },
    getPendingDeleteCount() {
      return entries.filter((e) => e.pending === "delete").length;
    },
    getItemsVisible() {
      return seriesItems != null;
    },
    /** Host Chart.name after load (distinct from lookup key "C1"). */
    setLoadedChartName(name: unknown) {
      chartNameValue = name;
    },
    /** delete → load (old committed) → single sync: items stay old names. */
    async brokenDeleteSkipFirstSync(seriesIndex1: number) {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      chart.series.getItemAt(seriesIndex1 - 1).delete();
      chart.series.load("items/name,items/chartType,items/smooth");
      await context.sync();
      return (seriesItems ?? []).map((s) => s.name);
    },
  };
}
