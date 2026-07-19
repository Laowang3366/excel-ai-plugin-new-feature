/** Sync-gated fake for Worksheet.pageLayout (paperSize + zoom scale/fit + OrNullObject). */
export function installPageLayoutExcel(options?: {
  excelApi19?: boolean;
  isSetSupportedThrows?: boolean;
  missingIsSetSupported?: boolean;
  /** When false, pageLayout has no paperSize member. */
  hasPaperSize?: boolean;
  /** When false, pageLayout has no zoom member. */
  hasZoom?: boolean;
}) {
  type Layout = {
    orientation: string;
    centerHorizontally: boolean;
    centerVertically: boolean;
    printGridlines: boolean;
    printHeadings: boolean;
    blackAndWhite: boolean;
    topMargin: number;
    bottomMargin: number;
    leftMargin: number;
    rightMargin: number;
    paperSize: string;
    /** Official PageLayoutZoomOptions.scale may be null (fit-to-pages). */
    zoomScale: number | null;
    fitToPagesWide: number | null;
    fitToPagesTall: number | null;
    printArea: string | null;
    printTitleRows: string | null;
    printTitleColumns: string | null;
  };

  type SheetState = {
    name: string;
    committed: Layout;
    pending: Partial<Layout> | undefined;
  };

  function defaults(): Layout {
    return {
      orientation: "Portrait",
      centerHorizontally: false,
      centerVertically: false,
      printGridlines: false,
      printHeadings: false,
      blackAndWhite: false,
      topMargin: 72,
      bottomMargin: 72,
      leftMargin: 72,
      rightMargin: 72,
      paperSize: "Letter",
      zoomScale: 100,
      fitToPagesWide: null,
      fitToPagesTall: null,
      printArea: null,
      printTitleRows: null,
      printTitleColumns: null,
    };
  }

  const excelApi19 = options?.excelApi19 !== false;
  const hasPaperSize = options?.hasPaperSize !== false;
  const hasZoom = options?.hasZoom !== false;
  let excelRunCalls = 0;
  let pageLayoutWriteCalls = 0;

  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", { name: "Sheet1", committed: defaults(), pending: undefined });
  sheets.set("Sheet2", { name: "Sheet2", committed: defaults(), pending: undefined });

  function queue(sheet: SheetState, patch: Partial<Layout>) {
    pageLayoutWriteCalls += 1;
    sheet.pending = { ...(sheet.pending ?? {}), ...patch };
  }

  /** RangeAreas-shaped OrNullObject (print area). */
  function rangeAreasOrNull(getAddress: () => string | null) {
    return {
      load() {},
      get isNullObject() {
        return getAddress() == null;
      },
      get address() {
        return getAddress() ?? "";
      },
    };
  }

  /** Range-shaped OrNullObject (print title rows/columns). */
  function rangeOrNull(getAddress: () => string | null) {
    return {
      load() {},
      get isNullObject() {
        return getAddress() == null;
      },
      get address() {
        return getAddress() ?? "";
      },
      values: [] as unknown[],
      formulas: [] as unknown[],
      rowCount: 0,
      columnCount: 0,
    };
  }

  function makePageLayout(sheet: SheetState) {
    const layout: Record<string, unknown> = {
      load() {},
      get orientation() {
        return sheet.committed.orientation;
      },
      set orientation(v: string) {
        queue(sheet, { orientation: v });
      },
      get centerHorizontally() {
        return sheet.committed.centerHorizontally;
      },
      set centerHorizontally(v: boolean) {
        queue(sheet, { centerHorizontally: v });
      },
      get centerVertically() {
        return sheet.committed.centerVertically;
      },
      set centerVertically(v: boolean) {
        queue(sheet, { centerVertically: v });
      },
      get printGridlines() {
        return sheet.committed.printGridlines;
      },
      set printGridlines(v: boolean) {
        queue(sheet, { printGridlines: v });
      },
      get printHeadings() {
        return sheet.committed.printHeadings;
      },
      set printHeadings(v: boolean) {
        queue(sheet, { printHeadings: v });
      },
      get blackAndWhite() {
        return sheet.committed.blackAndWhite;
      },
      set blackAndWhite(v: boolean) {
        queue(sheet, { blackAndWhite: v });
      },
      get topMargin() {
        return sheet.committed.topMargin;
      },
      set topMargin(v: number) {
        queue(sheet, { topMargin: v });
      },
      get bottomMargin() {
        return sheet.committed.bottomMargin;
      },
      set bottomMargin(v: number) {
        queue(sheet, { bottomMargin: v });
      },
      get leftMargin() {
        return sheet.committed.leftMargin;
      },
      set leftMargin(v: number) {
        queue(sheet, { leftMargin: v });
      },
      get rightMargin() {
        return sheet.committed.rightMargin;
      },
      set rightMargin(v: number) {
        queue(sheet, { rightMargin: v });
      },
      getPrintAreaOrNullObject() {
        return rangeAreasOrNull(() => sheet.committed.printArea);
      },
      setPrintArea(address: string) {
        queue(sheet, { printArea: address });
      },
      getPrintTitleRowsOrNullObject() {
        return rangeOrNull(() => sheet.committed.printTitleRows);
      },
      setPrintTitleRows(address: string) {
        queue(sheet, { printTitleRows: address });
      },
      getPrintTitleColumnsOrNullObject() {
        return rangeOrNull(() => sheet.committed.printTitleColumns);
      },
      setPrintTitleColumns(address: string) {
        queue(sheet, { printTitleColumns: address });
      },
    };

    if (hasPaperSize) {
      Object.defineProperty(layout, "paperSize", {
        enumerable: true,
        configurable: true,
        get() {
          return sheet.committed.paperSize;
        },
        set(v: string) {
          queue(sheet, { paperSize: v });
        },
      });
    }

    if (hasZoom) {
      Object.defineProperty(layout, "zoom", {
        enumerable: true,
        configurable: true,
        get() {
          return {
            scale: sheet.committed.zoomScale,
            horizontalFitToPages: sheet.committed.fitToPagesWide ?? undefined,
            verticalFitToPages: sheet.committed.fitToPagesTall ?? undefined,
          };
        },
        set(options: {
          scale?: number | null;
          horizontalFitToPages?: number;
          verticalFitToPages?: number;
        }) {
          if (!options || typeof options !== "object") return;
          const patch: Partial<Layout> = {};
          if (Object.prototype.hasOwnProperty.call(options, "scale")) {
            patch.zoomScale = options.scale ?? null;
            patch.fitToPagesWide = null;
            patch.fitToPagesTall = null;
          }
          if (Object.prototype.hasOwnProperty.call(options, "horizontalFitToPages")) {
            patch.fitToPagesWide = options.horizontalFitToPages ?? null;
            patch.zoomScale = null;
          }
          if (Object.prototype.hasOwnProperty.call(options, "verticalFitToPages")) {
            patch.fitToPagesTall = options.verticalFitToPages ?? null;
            patch.zoomScale = null;
          }
          if (Object.keys(patch).length > 0) queue(sheet, patch);
        },
      });
    }

    return layout;
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      load() {},
      pageLayout: makePageLayout(sheet),
    };
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          return makeSheet(name);
        },
      },
    },
    async sync() {
      for (const sheet of sheets.values()) {
        if (sheet.pending) {
          sheet.committed = { ...sheet.committed, ...sheet.pending };
          sheet.pending = undefined;
        }
      }
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => {
      excelRunCalls += 1;
      return fn(context);
    },
  };

  if (options?.missingIsSetSupported) {
    (globalThis as unknown as { Office?: unknown }).Office = { context: { requirements: {} } };
  } else if (options?.isSetSupportedThrows) {
    (globalThis as unknown as {
      Office: { context: { requirements: { isSetSupported: () => boolean } } };
    }).Office = {
      context: {
        requirements: {
          isSetSupported() {
            throw new Error("isSetSupported threw");
          },
        },
      },
    };
  } else {
    (globalThis as unknown as {
      Office: {
        context: { requirements: { isSetSupported: (n: string, v?: string) => boolean } };
      };
    }).Office = {
      context: {
        requirements: {
          isSetSupported(name: string, minVersion?: string) {
            if (name === "ExcelApi" && minVersion === "1.9") return excelApi19;
            return false;
          },
        },
      },
    };
  }

  return {
    getCommitted(name: string) {
      return sheets.get(name)?.committed;
    },
    getPending(name: string) {
      return sheets.get(name)?.pending;
    },
    getExcelRunCalls() {
      return excelRunCalls;
    },
    getPageLayoutWriteCalls() {
      return pageLayoutWriteCalls;
    },
    /** Seed committed zoom.scale (including null) for null-readback tests. */
    setCommittedZoomScale(name: string, scale: number | null) {
      const sheet = sheets.get(name);
      if (sheet) {
        sheet.committed.zoomScale = scale;
        if (scale != null) {
          sheet.committed.fitToPagesWide = null;
          sheet.committed.fitToPagesTall = null;
        }
      }
    },
    setCommittedPaperSize(name: string, paperSize: string) {
      const sheet = sheets.get(name);
      if (sheet) sheet.committed.paperSize = paperSize;
    },
    /** Change host worksheet.name without changing getItem lookup key (anti input-echo). */
    setHostSheetName(lookupName: string, hostName: string) {
      const sheet = sheets.get(lookupName);
      if (sheet) sheet.name = hostName;
    },
    setCommittedFit(
      name: string,
      wide: number | null,
      tall: number | null,
    ) {
      const sheet = sheets.get(name);
      if (sheet) {
        sheet.committed.fitToPagesWide = wide;
        sheet.committed.fitToPagesTall = tall;
        if (wide != null || tall != null) sheet.committed.zoomScale = null;
      }
    },
  };
}
