/** Sync-gated fake for Worksheet.pageLayout (print scalars + paperSize/zoom + OrNullObject). */
import {
  defaultPageLayoutState,
  makePageLayoutObject,
  type PageLayoutSheetState,
  type PageLayoutState,
} from "./officeJsPageLayoutFakeLayout";

export function installPageLayoutExcel(options?: {
  excelApi19?: boolean;
  isSetSupportedThrows?: boolean;
  missingIsSetSupported?: boolean;
  /** When false, pageLayout has no paperSize member. */
  hasPaperSize?: boolean;
  /** When false, pageLayout has no zoom member. */
  hasZoom?: boolean;
  /** When false, pageLayout has no draftMode member. */
  hasDraftMode?: boolean;
  /** When false, pageLayout has no printOrder member. */
  hasPrintOrder?: boolean;
  /** When false, pageLayout has no firstPageNumber member. */
  hasFirstPageNumber?: boolean;
  /** When false, pageLayout has no headerMargin member. */
  hasHeaderMargin?: boolean;
  /** When false, pageLayout has no footerMargin member. */
  hasFooterMargin?: boolean;
  /** When false, pageLayout has no headersFooters member. */
  hasHeadersFooters?: boolean;
  /** When false, headersFooters has no defaultForAllPages. */
  hasDefaultForAllPages?: boolean;
  /** Omit one defaultForAllPages slot property. */
  missingHeaderFooterSlot?:
    | "leftHeader"
    | "centerHeader"
    | "rightHeader"
    | "leftFooter"
    | "centerFooter"
    | "rightFooter";
}) {
  const excelApi19 = options?.excelApi19 !== false;
  const hasPaperSize = options?.hasPaperSize !== false;
  const hasZoom = options?.hasZoom !== false;
  const hasDraftMode = options?.hasDraftMode !== false;
  const hasPrintOrder = options?.hasPrintOrder !== false;
  const hasFirstPageNumber = options?.hasFirstPageNumber !== false;
  const hasHeaderMargin = options?.hasHeaderMargin !== false;
  const hasFooterMargin = options?.hasFooterMargin !== false;
  const hasHeadersFooters = options?.hasHeadersFooters !== false;
  const hasDefaultForAllPages = options?.hasDefaultForAllPages !== false;
  let excelRunCalls = 0;
  let pageLayoutWriteCalls = 0;

  const sheets = new Map<string, PageLayoutSheetState>();
  sheets.set("Sheet1", {
    name: "Sheet1",
    committed: defaultPageLayoutState(),
    pending: undefined,
  });
  sheets.set("Sheet2", {
    name: "Sheet2",
    committed: defaultPageLayoutState(),
    pending: undefined,
  });

  function queue(sheet: PageLayoutSheetState, patch: Partial<PageLayoutState>) {
    pageLayoutWriteCalls += 1;
    sheet.pending = { ...(sheet.pending ?? {}), ...patch };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      load() {},
      pageLayout: makePageLayoutObject(sheet, {
        hasPaperSize,
        hasZoom,
        hasDraftMode,
        hasPrintOrder,
        hasFirstPageNumber,
        hasHeaderMargin,
        hasFooterMargin,
        hasHeadersFooters,
        hasDefaultForAllPages,
        missingHeaderFooterSlot: options?.missingHeaderFooterSlot,
        queue,
      }),
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
    setCommittedFirstPageNumber(name: string, value: number | string | null) {
      const sheet = sheets.get(name);
      if (sheet) sheet.committed.firstPageNumber = value;
    },
    setCommittedPrintOrder(name: string, value: string) {
      const sheet = sheets.get(name);
      if (sheet) sheet.committed.printOrder = value;
    },
    /** Change host worksheet.name without changing getItem lookup key (anti input-echo). */
    setHostSheetName(lookupName: string, hostName: unknown) {
      const sheet = sheets.get(lookupName);
      if (sheet) sheet.name = hostName;
    },
    setCommittedFit(name: string, wide: number | null, tall: number | null) {
      const sheet = sheets.get(name);
      if (sheet) {
        sheet.committed.fitToPagesWide = wide;
        sheet.committed.fitToPagesTall = tall;
        if (wide != null || tall != null) sheet.committed.zoomScale = null;
      }
    },
    setCommittedHeaderFooterMargins(name: string, header: number, footer: number) {
      const sheet = sheets.get(name);
      if (sheet) {
        sheet.committed.headerMargin = header;
        sheet.committed.footerMargin = footer;
      }
    },
    setCommittedHeadersFooters(
      name: string,
      values: Partial<{
        leftHeader: string;
        centerHeader: string;
        rightHeader: string;
        leftFooter: string;
        centerFooter: string;
        rightFooter: string;
      }>,
    ) {
      const sheet = sheets.get(name);
      if (!sheet) return;
      Object.assign(sheet.committed, values);
    },
  };
}
