/**
 * Sync-gated fake for ChartSeriesCollection.add().
 * add() only marks pending; committed collection grows on first context.sync().
 * load() snapshots committed-at-load; skip-add-sync (add→load→sync) keeps old items.
 */

type SeriesState = { name: string; chartType: string; smooth: boolean };

type Entry = {
  committed: SeriesState | null;
  pending: SeriesState | undefined;
  order: number;
};

export function installChartSeriesAddExcel(initial: SeriesState[] = []) {
  const entries: Entry[] = initial.map((s, i) => ({
    committed: { ...s },
    pending: undefined,
    order: i,
  }));
  let nextOrder = initial.length;
  let nextDefault = 1;

  let seriesItems: SeriesState[] | null = null;
  let pendingSeriesItems: SeriesState[] | null = null;
  let chartNameValue: unknown = "C1";
  let nextAddChartType = "ColumnClustered";

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
                    add(name?: string) {
                      const resolved =
                        name != null && name !== ""
                          ? name
                          : `Series${nextDefault++}`;
                      entries.push({
                        committed: null,
                        pending: {
                          name: resolved,
                          chartType: nextAddChartType,
                          smooth: false,
                        },
                        order: nextOrder++,
                      });
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
        if (entry.pending != null) {
          entry.committed = entry.pending;
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
    setLoadedChartName(name: unknown) {
      chartNameValue = name;
    },
    setNextAddChartType(chartType: string) {
      nextAddChartType = chartType;
    },
    /** add → load (old committed) → single sync: items stay old. */
    async brokenAddSkipFirstSync(name?: string) {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      if (name != null) chart.series.add(name);
      else chart.series.add();
      chart.series.load("items/name,items/chartType,items/smooth");
      await context.sync();
      return (seriesItems ?? []).map((s) => s.name);
    },
  };
}
