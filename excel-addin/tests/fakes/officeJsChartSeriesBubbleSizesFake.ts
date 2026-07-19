/**
 * Sync-gated fake for ChartSeries.setBubbleSizes +
 * getDimensionDataSourceString("BubbleSizes") (ExcelApi 1.15).
 * Office.context.requirements.isSetSupported precheck must pass before setters.
 */

type SeriesState = {
  bubbleSizesSource: string | null;
  pendingBubbleSizes: string | null | undefined;
};

export function installChartSeriesBubbleSizesExcel(options?: {
  seriesCount?: number;
  supportReadback?: boolean;
  hostSourcePrefix?: string;
  failReadbackSync?: boolean;
  /** When false, isSetSupported('ExcelApi','1.15') returns false. Default true. */
  excelApi115?: boolean;
  /** Override readback payload after load flush (empty / non-string for failure tests). */
  readbackValue?: unknown;
  sheetNameValue?: unknown;
  chartNameValue?: unknown;
}) {
  const seriesCount = options?.seriesCount ?? 2;
  const supportReadback = options?.supportReadback !== false;
  const hostSourcePrefix = options?.hostSourcePrefix ?? "Sheet1!";
  const failReadbackSync = options?.failReadbackSync === true;
  const excelApi115 = options?.excelApi115 !== false;
  const seriesList: SeriesState[] = Array.from({ length: seriesCount }, () => ({
    bubbleSizesSource: null,
    pendingBubbleSizes: undefined,
  }));

  let sheetNameValue: unknown = options?.sheetNameValue ?? "Sheet1";
  let chartNameValue: unknown = options?.chartNameValue ?? "C1";
  let loadedSource: unknown | undefined = undefined;
  let sourceLoaded = false;
  let pendingLoad: unknown | undefined = undefined;
  let hasPendingLoad = false;
  let snapshot: Array<{ bubbleSizesSource: string | null }> | null = null;
  let setBubbleSizesCalls = 0;

  function hostSourceString(bare: string): string {
    return `${hostSourcePrefix}${bare}`;
  }

  function captureSnapshot() {
    if (snapshot) return;
    snapshot = seriesList.map((s) => ({
      bubbleSizesSource: s.bubbleSizesSource,
    }));
  }

  function restoreSnapshot() {
    if (!snapshot) return;
    for (let i = 0; i < seriesList.length; i += 1) {
      const s = seriesList[i]!;
      const prev = snapshot[i]!;
      s.bubbleSizesSource = prev.bubbleSizesSource;
      s.pendingBubbleSizes = undefined;
    }
    snapshot = null;
    pendingLoad = undefined;
    hasPendingLoad = false;
    loadedSource = undefined;
    sourceLoaded = false;
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            get name() {
              return sheetNameValue as string;
            },
            load() {},
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
                        setBubbleSizes(range: { address: string }): void;
                        getDimensionDataSourceString?: (dim: "BubbleSizes") => { value: unknown };
                      } = {
                        setBubbleSizes(range) {
                          setBubbleSizesCalls += 1;
                          captureSnapshot();
                          state.pendingBubbleSizes = hostSourceString(range.address);
                        },
                      };
                      if (supportReadback) {
                        series.getDimensionDataSourceString = (dim: "BubbleSizes") => {
                          if (dim !== "BubbleSizes") {
                            throw new Error(`unexpected dimension ${dim}`);
                          }
                          hasPendingLoad = true;
                          pendingLoad =
                            options?.readbackValue !== undefined
                              ? options.readbackValue
                              : state.bubbleSizesSource;
                          return {
                            get value() {
                              if (!sourceLoaded) {
                                throw new Error("source not loaded before sync");
                              }
                              return loadedSource;
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
      const isReadbackFlush = hasPendingLoad;
      for (const state of seriesList) {
        if (state.pendingBubbleSizes !== undefined) {
          state.bubbleSizesSource = state.pendingBubbleSizes;
          state.pendingBubbleSizes = undefined;
        }
      }
      if (isReadbackFlush && failReadbackSync) {
        restoreSnapshot();
        throw new Error("bubble size source readback sync rejected");
      }
      if (hasPendingLoad) {
        loadedSource = pendingLoad;
        sourceLoaded = true;
        pendingLoad = undefined;
        hasPendingLoad = false;
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
      return { setBubbleSizesCalls };
    },
    setLoadedSheetName(name: unknown) {
      sheetNameValue = name;
    },
    setLoadedChartName(name: unknown) {
      chartNameValue = name;
    },
    /**
     * Stale path: setBubbleSizes(B) then getDimension before write flush samples
     * committed A (or null); one sync commits B but readback value stays stale A/null.
     */
    async brokenSkipFirstSync(seriesIndex1: number, bare: string) {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      const series = chart.series.getItemAt(seriesIndex1 - 1);
      series.setBubbleSizes({ address: bare });
      const r = series.getDimensionDataSourceString!("BubbleSizes");
      await context.sync();
      try {
        return r.value;
      } catch {
        return null;
      }
    },
  };
}
