/**
 * Sync-gated fake for Range.getImage().
 * ClientResult.value is readable only after context.sync().
 * Host sheet name / address may differ from request keys.
 */

export function installRangeImageExcel(options?: {
  excelApi17?: boolean;
  imagePayload?: string | null;
  /** Host worksheet.name after load (may differ from request key). */
  hostSheetName?: unknown;
  /** Host Range.address after load. */
  hostAddress?: unknown;
  /** When true, range has no getImage function. */
  missingGetImage?: boolean;
  syncFails?: boolean;
}) {
  const excelApi17 = options?.excelApi17 !== false;
  let hostSheetName: unknown =
    options && "hostSheetName" in options ? options.hostSheetName : "HostSheet";
  let hostAddress: unknown =
    options && "hostAddress" in options ? options.hostAddress : "HostSheet!$A$1:$B$2";
  let imagePayload: string | null =
    options && "imagePayload" in options ? (options.imagePayload as string | null) : "aG9zdC1yYW5nZS1iYXNlNjQ=";
  const missingGetImage = options?.missingGetImage === true;
  const syncFails = options?.syncFails === true;
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
            getRange(address: string) {
              if (address !== "A1:B2" && address !== "A1") {
                throw new Error(`missing range ${address}`);
              }
              const range: {
                get address(): unknown;
                load(): void;
                getImage?: () => { get value(): string };
              } = {
                get address() {
                  return hostAddress;
                },
                load() {},
              };
              if (!missingGetImage) {
                range.getImage = () => {
                  getImageCalls += 1;
                  pendingImage = imagePayload;
                  return {
                    get value() {
                      if (committedImage == null) {
                        throw new Error("ClientResult.value not available before sync");
                      }
                      return committedImage;
                    },
                  };
                };
              }
              return range;
            },
          };
        },
      },
    },
    async sync() {
      if (syncFails) throw new Error("sync failed");
      committedImage = pendingImage;
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
          if (name === "ExcelApi" && minVersion === "1.7") return excelApi17;
          return false;
        },
      },
    },
  };
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getImageCalls: () => getImageCalls,
    setImagePayload(value: string | null) {
      imagePayload = value;
    },
    setHostSheetName(value: unknown) {
      hostSheetName = value;
    },
    setHostAddress(value: unknown) {
      hostAddress = value;
    },
    async brokenSkipSync() {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const range = sheet.getRange("A1:B2");
      if (typeof range.getImage !== "function") return null;
      const result = range.getImage();
      try {
        return result.value;
      } catch {
        return null;
      }
    },
    async staleAfterPayloadChange() {
      const sheet = context.workbook.worksheets.getItem("Sheet1");
      const range = sheet.getRange("A1:B2");
      if (typeof range.getImage !== "function") {
        throw new Error("getImage missing");
      }
      const firstResult = range.getImage();
      await context.sync();
      const first = firstResult.value;
      imagePayload = "cGF5bG9hZEI=";
      const secondResult = range.getImage();
      const stale = secondResult.value;
      return { first, stale, pending: pendingImage };
    },
  };
}
