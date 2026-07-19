/**
 * Sync-gated fake for ChartSeries.setValues/setXAxisValues +
 * getDimensionDataSourceString (ExcelApi 1.15).
 * Office.context.requirements.isSetSupported precheck must pass before setters.
 */

type Dim = "Values" | "XValues";

type SeriesState = {
  valuesSource: string | null;
  xValuesSource: string | null;
  pendingValues: string | null | undefined;
  pendingXValues: string | null | undefined;
};

export function installChartSeriesValuesExcel(options?: {
  seriesCount?: number;
  supportReadback?: boolean;
  hostSourcePrefix?: string;
  failReadbackSync?: boolean;
  /** When false, isSetSupported('ExcelApi','1.15') returns false. Default true. */
  excelApi115?: boolean;
}) {
  const seriesCount = options?.seriesCount ?? 2;
  const supportReadback = options?.supportReadback !== false;
  const hostSourcePrefix = options?.hostSourcePrefix ?? "Sheet1!";
  const failReadbackSync = options?.failReadbackSync === true;
  const excelApi115 = options?.excelApi115 !== false;
  const seriesList: SeriesState[] = Array.from({ length: seriesCount }, () => ({
    valuesSource: null,
    xValuesSource: null,
    pendingValues: undefined,
    pendingXValues: undefined,
  }));

  let chartNameValue: unknown = "C1";
  let loadedSources: Partial<Record<Dim, string | null>> | null = null;
  let pendingLoad: Partial<Record<Dim, string | null>> | null = null;
  let snapshot: Array<{ valuesSource: string | null; xValuesSource: string | null }> | null =
    null;
  let setValuesCalls = 0;
  let setXAxisValuesCalls = 0;

  function hostSourceString(bare: string): string {
    return `${hostSourcePrefix}${bare}`;
  }

  function captureSnapshot() {
    if (snapshot) return;
    snapshot = seriesList.map((s) => ({
      valuesSource: s.valuesSource,
      xValuesSource: s.xValuesSource,
    }));
  }

  function restoreSnapshot() {
    if (!snapshot) return;
    for (let i = 0; i < seriesList.length; i += 1) {
      const s = seriesList[i]!;
      const prev = snapshot[i]!;
      s.valuesSource = prev.valuesSource;
      s.xValuesSource = prev.xValuesSource;
      s.pendingValues = undefined;
      s.pendingXValues = undefined;
    }
    snapshot = null;
    pendingLoad = null;
    loadedSources = null;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            getRange(address: string) {
              return { address };
            },
            charts: {
              getItem(chartName: string) {
                if (chartName !== "C1") throw new Error(`missing chart ${chartName}`);
                return {
                  get name() {
                    return chartNameValue as string;
                  },
                  load() {},
                  series: {
                    getItemAt(index: number) {
                      if (index < 0 || index >= seriesList.length) {
                        throw new Error(`series index out of range: ${index}`);
                      }
                      const state = seriesList[index]!;
                      const series: {
                        setValues(range: { address: string }): void;
                        setXAxisValues(range: { address: string }): void;
                        getDimensionDataSourceString?: (dim: Dim) => { value: string };
                      } = {
                        setValues(range) {
                          setValuesCalls += 1;
                          captureSnapshot();
                          state.pendingValues = hostSourceString(range.address);
                        },
                        setXAxisValues(range) {
                          setXAxisValuesCalls += 1;
                          captureSnapshot();
                          state.pendingXValues = hostSourceString(range.address);
                        },
                      };
                      if (supportReadback) {
                        series.getDimensionDataSourceString = (dim: Dim) => {
                          pendingLoad = pendingLoad ?? {};
                          if (dim === "Values") pendingLoad.Values = state.valuesSource;
                          else pendingLoad.XValues = state.xValuesSource;
                          return {
                            get value() {
                              if (!loadedSources) {
                                throw new Error("source not loaded before sync");
                              }
                              const v = loadedSources[dim];
                              if (v == null) throw new Error(`missing source ${dim}`);
                              return v;
                            },
                          };
                        };
                      }
                      return series;
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
      const isReadbackFlush = pendingLoad != null;
      for (const state of seriesList) {
        if (state.pendingValues !== undefined) {
          state.valuesSource = state.pendingValues;
          state.pendingValues = undefined;
        }
        if (state.pendingXValues !== undefined) {
          state.xValuesSource = state.pendingXValues;
          state.pendingXValues = undefined;
        }
      }
      if (isReadbackFlush && failReadbackSync) {
        restoreSnapshot();
        throw new Error(
          "getDimensionDataSourceString readback sync rejected (ExcelApi 1.15 required)",
        );
      }
      if (pendingLoad) {
        loadedSources = pendingLoad;
        pendingLoad = null;
        snapshot = null;
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as {
    Office: {
      context: {
        requirements: { isSetSupported: (name: string, minVersion?: string) => boolean };
      };
    };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(name: string, minVersion?: string) {
          if (name === "ExcelApi" && minVersion === "1.15") return excelApi115;
          return false;
        },
      },
    },
  };
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getCommitted(index0: number) {
      return seriesList[index0];
    },
    getSetterCallCounts() {
      return { setValuesCalls, setXAxisValuesCalls };
    },
    setLoadedChartName(name: unknown) {
      chartNameValue = name;
    },
    async brokenSkipFirstSync(seriesIndex1: number, valuesBare: string) {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      const series = chart.series.getItemAt(seriesIndex1 - 1);
      series.setValues({ address: valuesBare });
      const r = series.getDimensionDataSourceString!("Values");
      await context.sync();
      try {
        return r.value;
      } catch {
        return null;
      }
    },
  };
}
