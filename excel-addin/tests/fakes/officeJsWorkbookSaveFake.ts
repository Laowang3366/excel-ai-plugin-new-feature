/** Fake Excel.run for workbook.save with sync-gated save + optional failures. */
export type WorkbookSaveFakeOptions = {
  excelApi11?: boolean;
  /** Throw from isSetSupported (fail-safe → unsupported). */
  isSetSupportedThrows?: boolean;
  /** Missing requirements.isSetSupported. */
  missingIsSetSupported?: boolean;
  /** Throw during context.sync after save queued. */
  syncError?: string | null;
  workbookName?: string;
};

export function installWorkbookSaveExcel(options: WorkbookSaveFakeOptions = {}) {
  const {
    excelApi11 = true,
    isSetSupportedThrows = false,
    missingIsSetSupported = false,
    syncError = null,
    workbookName = "Book1.xlsx",
  } = options;

  let saveQueued = false;
  let syncCount = 0;
  let saveCallCount = 0;

  const context = {
    workbook: {
      name: workbookName,
      load() {},
      save() {
        saveCallCount += 1;
        saveQueued = true;
      },
    },
    async sync() {
      syncCount += 1;
      if (syncError) throw new Error(syncError);
      if (!saveQueued) throw new Error("sync without save queued");
      saveQueued = false;
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  const requirements = missingIsSetSupported
    ? {}
    : {
        isSetSupported(name: string, version: string) {
          if (isSetSupportedThrows) throw new Error("isSetSupported boom");
          if (name !== "ExcelApi") return false;
          return excelApi11 && version === "1.1";
        },
      };

  (globalThis as unknown as { Office: unknown }).Office = {
    context: { requirements },
  };

  return {
    context,
    saveCallCount: () => saveCallCount,
    syncCount: () => syncCount,
  };
}

export function uninstallWorkbookSaveExcel() {
  const g = globalThis as unknown as {
    Excel?: unknown;
    Office?: unknown;
    window?: unknown;
  };
  delete g.Excel;
  delete g.Office;
}
