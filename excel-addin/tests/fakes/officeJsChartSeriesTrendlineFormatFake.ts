/**
 * Sync-gated fake for ChartTrendline.format.line (ExcelApi 1.7).
 */

type LineState = {
  color: string;
  lineStyle: string;
  weight: number;
};

type LineEntry = {
  committed: LineState;
  pending?: Partial<LineState>;
  loadQueued?: boolean;
  snapshot?: LineState;
};

export function installChartSeriesTrendlineFormatExcel(options?: {
  excelApi17?: boolean;
  seriesCount?: number;
  trendlineCount?: number;
}): {
  getCommitted(series0: number, trendline0: number): LineState | undefined;
  poison(series0: number, trendline0: number, patch: Partial<LineState>): void;
  coerceStyleAfterWrite(hostStyle: string | null): void;
  setChartName(v: unknown): void;
} {
  const excelApi17 = options?.excelApi17 !== false;
  const seriesCount = options?.seriesCount ?? 1;
  const trendlineCount = options?.trendlineCount ?? 2;
  let chartNameValue: unknown = "HostChart";
  let coerceStyle: string | null = null;

  const lines: LineEntry[][] = Array.from({ length: seriesCount }, () =>
    Array.from({ length: trendlineCount }, () => ({
      committed: {
        color: "#000000",
        lineStyle: "Continuous",
        weight: 1.5,
      },
    })),
  );

  function makeLineProxy(entry: LineEntry) {
    return {
      get color() {
        if (!entry.snapshot) throw new Error("ChartLineFormat.color not loaded");
        return entry.snapshot.color;
      },
      set color(v: string) {
        entry.pending = { ...entry.pending, color: v };
      },
      get lineStyle() {
        if (!entry.snapshot) throw new Error("ChartLineFormat.lineStyle not loaded");
        return entry.snapshot.lineStyle;
      },
      set lineStyle(v: string) {
        entry.pending = { ...entry.pending, lineStyle: v };
      },
      get weight() {
        if (!entry.snapshot) throw new Error("ChartLineFormat.weight not loaded");
        return entry.snapshot.weight;
      },
      set weight(v: number) {
        entry.pending = { ...entry.pending, weight: v };
      },
      load(_props: string) {
        entry.loadQueued = true;
      },
    };
  }

  const proxies = lines.map((series) => series.map((e) => makeLineProxy(e)));

  const context = {
    workbook: {
      worksheets: {
        getItem(_name: string) {
          return {
            charts: {
              getItem(_chartName: string) {
                return {
                  get name() {
                    return chartNameValue;
                  },
                  load() {},
                  series: {
                    getItemAt(seriesIndex: number) {
                      if (seriesIndex < 0 || seriesIndex >= proxies.length) {
                        throw new Error(`series index out of range: ${seriesIndex}`);
                      }
                      return {
                        trendlines: {
                          getItem(trendlineIndex: number) {
                            if (
                              trendlineIndex < 0 ||
                              trendlineIndex >= proxies[seriesIndex]!.length
                            ) {
                              throw new Error(`trendline index out of range: ${trendlineIndex}`);
                            }
                            return {
                              format: {
                                line: proxies[seriesIndex]![trendlineIndex]!,
                              },
                            };
                          },
                        },
                      };
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
      for (const series of lines) {
        for (const entry of series) {
          if (entry.pending) {
            entry.committed = { ...entry.committed, ...entry.pending };
            entry.pending = undefined;
            if (coerceStyle != null) {
              entry.committed.lineStyle = coerceStyle;
              coerceStyle = null;
            }
          }
        }
      }
      for (const series of lines) {
        for (const entry of series) {
          if (entry.loadQueued) {
            entry.snapshot = { ...entry.committed };
            entry.loadQueued = false;
          }
        }
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };
  (globalThis as unknown as {
    Office: {
      context: { requirements: { isSetSupported: (n: string, v?: string) => boolean } };
    };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(name: string, minVersion?: string) {
          if (name !== "ExcelApi") return false;
          if (minVersion === "1.7") return excelApi17;
          return true;
        },
      },
    },
  };

  return {
    getCommitted(series0: number, trendline0: number) {
      const e = lines[series0]?.[trendline0];
      return e ? { ...e.committed } : undefined;
    },
    poison(series0: number, trendline0: number, patch: Partial<LineState>) {
      const e = lines[series0]?.[trendline0];
      if (!e) throw new Error("no trendline");
      e.committed = { ...e.committed, ...patch };
    },
    coerceStyleAfterWrite(hostStyle: string | null) {
      coerceStyle = hostStyle;
    },
    setChartName(v: unknown) {
      chartNameValue = v;
    },
  };
}
