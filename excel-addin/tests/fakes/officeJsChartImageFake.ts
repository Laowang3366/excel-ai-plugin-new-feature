/**
 * Sync-gated fake for Chart.getImage().
 * ClientResult.value is readable only after context.sync().
 * Lookup key is request sheet/chart name; host name props may differ.
 */

export function installChartImageExcel(options?: {
  excelApi12?: boolean;
  imagePayload?: string;
  /** Host worksheet.name after load (may differ from request key). */
  hostSheetName?: string;
  chartName?: unknown;
}) {
  const excelApi12 = options?.excelApi12 !== false;
  let hostSheetName = options?.hostSheetName ?? "HostSheet";
  let chartNameValue: unknown = options?.chartName ?? "HostChart";
  let imagePayload = options?.imagePayload ?? "ZmFrZS1iYXNlNjQ=";
  let pendingImage: string | null = null;
  let committedImage: string | null = null;
  let getImageCalls = 0;

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          if (name !== "Sheet1") throw new Error(`missing sheet ${name}`);
          return {
            get name() {
              return hostSheetName;
            },
            load() {},
            charts: {
              getItem(chartName: string) {
                if (chartName !== "C1") throw new Error(`missing chart ${chartName}`);
                return {
                  get name() {
                    return chartNameValue as string;
                  },
                  load() {},
                  getImage(width?: number, height?: number) {
                    getImageCalls += 1;
                    const suffix =
                      width != null || height != null ? `:${width ?? ""}x${height ?? ""}` : "";
                    pendingImage = `${imagePayload}${suffix}`;
                    return {
                      get value() {
                        if (committedImage == null) {
                          throw new Error("image not available before sync");
                        }
                        return committedImage;
                      },
                    };
                  },
                };
              },
            },
          };
        },
      },
    },
    async sync() {
      if (pendingImage != null) {
        committedImage = pendingImage;
        pendingImage = null;
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
          if (name === "ExcelApi" && minVersion === "1.2") return excelApi12;
          return false;
        },
      },
    },
  };
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getImageCalls() {
      return getImageCalls;
    },
    setLoadedChartName(name: unknown) {
      chartNameValue = name;
    },
    setHostSheetName(name: string) {
      hostSheetName = name;
    },
    setImagePayload(payload: string) {
      imagePayload = payload;
    },
    getCommittedImage() {
      return committedImage;
    },
    async brokenSkipSync() {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      const r = chart.getImage();
      try {
        return r.value;
      } catch {
        return null;
      }
    },
    /** Commit A, change payload to B, getImage without sync → still A. */
    async staleAfterPayloadChange() {
      const chart = context.workbook.worksheets.getItem("Sheet1").charts.getItem("C1");
      chart.getImage();
      await context.sync();
      const first = committedImage;
      imagePayload = "cGF5bG9hZEI=";
      const r = chart.getImage();
      // no sync — value still first
      return { first, stale: r.value, pending: pendingImage };
    },
  };
}
