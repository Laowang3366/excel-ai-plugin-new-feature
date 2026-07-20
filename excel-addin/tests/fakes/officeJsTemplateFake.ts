/**
 * Sync-gated fake for workbook.template.apply / capture (UsedRange format + freeze + gridlines + pageLayout).
 */
import {
  defaultPageLayoutState,
  makePageLayoutObject,
  type PageLayoutSheetState,
} from "./officeJsPageLayoutFakeLayout";
import {
  attachPageBreaks,
  commitPageBreaks,
  defaultPageBreakSheetState,
  type PageBreakSheetState,
} from "./officeJsPageBreaksFake";

type FormatState = {
  font: { name: string; size: number; bold: boolean; color: string };
  fill: { color: string };
  horizontalAlignment: string;
  wrapText: boolean;
  rowHeight: number;
  columnWidth: number;
};

type SheetState = {
  name: string;
  empty: boolean;
  showGridlines: boolean;
  freezeRows: number;
  usedAddress: string;
  rows: number;
  cols: number;
  text: string;
  usedFormat: FormatState;
  headerFormat: FormatState;
  pageLayout: PageLayoutSheetState;
  pageBreaks: PageBreakSheetState;
  /** Poison overrides after write for readback failure tests. */
  poison?: Partial<{
    fontName: unknown;
    fontSize: unknown;
    headerFill: unknown;
    headerFontColor: unknown;
    headerBold: unknown;
    headerAlignment: unknown;
    headerWrap: unknown;
    headerRowHeight: unknown;
    showGridlines: unknown;
    freezeRows: unknown;
    address: unknown;
    rowCount: unknown;
    columnCount: unknown;
  }>;
};

function defaultFormat(): FormatState {
  return {
    font: { name: "Calibri", size: 11, bold: false, color: "#000000" },
    fill: { color: "#FFFFFF" },
    horizontalAlignment: "General",
    wrapText: false,
    rowHeight: 15,
    columnWidth: 10,
  };
}

export function installTemplateExcel(options?: {
  excelApi18?: boolean;
  excelApi19?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  missingShowGridlines?: boolean;
  missingFreeze?: boolean;
  missingAutofit?: boolean;
  missingUsedRange?: boolean;
  sheetCount?: number;
  activeSheet?: string;
}) {
  const excelApi18 = options?.excelApi18 !== false;
  const excelApi19 = options?.excelApi19 !== false;
  let excelRunCalls = 0;
  let writeCalls = 0;

  const sheets = new Map<string, SheetState>();
  function seed(name: string, empty = false) {
    sheets.set(name, {
      name,
      empty,
      showGridlines: true,
      freezeRows: 0,
      usedAddress: empty ? "A1" : "A1:B2",
      rows: empty ? 1 : 2,
      cols: empty ? 1 : 2,
      text: empty ? "" : "x",
      usedFormat: defaultFormat(),
      headerFormat: defaultFormat(),
      pageLayout: { name, committed: defaultPageLayoutState(), pending: undefined },
      pageBreaks: defaultPageBreakSheetState(),
    });
  }
  const count = options?.sheetCount ?? 2;
  if (count > 500) {
    for (let i = 1; i <= count; i += 1) seed(`S${i}`);
  } else {
    seed("Sheet1", false);
    seed("Sheet2", false);
    seed("Empty", true);
    for (let i = 3; i <= count; i += 1) seed(`Sheet${i}`, false);
  }

  const activeName = options?.activeSheet ?? "Sheet1";

  function makeRange(sheet: SheetState, kind: "used" | "header" | "other", address: string) {
    const formatState = kind === "header" ? sheet.headerFormat : sheet.usedFormat;
    let fontLoaded = false;
    let fillLoaded = false;
    let formatLoaded = false;
    const format: Record<string, unknown> = {
      get font() {
        return {
          get name() {
            if (!fontLoaded) throw new Error("font.name not loaded");
            if (sheet.poison?.fontName !== undefined && kind === "used") return sheet.poison.fontName as string;
            return formatState.font.name;
          },
          set name(v: string) {
            writeCalls += 1;
            formatState.font.name = v;
          },
          get size() {
            if (!fontLoaded) throw new Error("font.size not loaded");
            if (sheet.poison?.fontSize !== undefined && kind === "used") return sheet.poison.fontSize as number;
            return formatState.font.size;
          },
          set size(v: number) {
            writeCalls += 1;
            formatState.font.size = v;
          },
          get bold() {
            if (!fontLoaded) throw new Error("font.bold not loaded");
            if (sheet.poison?.headerBold !== undefined && kind === "header") {
              return sheet.poison.headerBold as boolean;
            }
            return formatState.font.bold;
          },
          set bold(v: boolean) {
            writeCalls += 1;
            formatState.font.bold = v;
          },
          get color() {
            if (!fontLoaded) throw new Error("font.color not loaded");
            if (sheet.poison?.headerFontColor !== undefined && kind === "header") {
              return sheet.poison.headerFontColor as string;
            }
            return formatState.font.color;
          },
          set color(v: string) {
            writeCalls += 1;
            formatState.font.color = v;
          },
          load() {
            fontLoaded = true;
          },
        };
      },
      get fill() {
        return {
          get color() {
            if (!fillLoaded) throw new Error("fill.color not loaded");
            if (sheet.poison?.headerFill !== undefined && kind === "header") {
              return sheet.poison.headerFill as string;
            }
            return formatState.fill.color;
          },
          set color(v: string) {
            writeCalls += 1;
            formatState.fill.color = v;
          },
          load() {
            fillLoaded = true;
          },
        };
      },
      get horizontalAlignment() {
        if (!formatLoaded) throw new Error("horizontalAlignment not loaded");
        if (sheet.poison?.headerAlignment !== undefined && kind === "header") {
          return sheet.poison.headerAlignment as string;
        }
        return formatState.horizontalAlignment;
      },
      set horizontalAlignment(v: string) {
        writeCalls += 1;
        formatState.horizontalAlignment = v;
      },
      get wrapText() {
        if (!formatLoaded) throw new Error("wrapText not loaded");
        if (sheet.poison?.headerWrap !== undefined && kind === "header") {
          return sheet.poison.headerWrap as boolean;
        }
        return formatState.wrapText;
      },
      set wrapText(v: boolean) {
        writeCalls += 1;
        formatState.wrapText = v;
      },
      get rowHeight() {
        if (!formatLoaded) throw new Error("rowHeight not loaded");
        if (sheet.poison?.headerRowHeight !== undefined && kind === "header") {
          return sheet.poison.headerRowHeight as number;
        }
        return formatState.rowHeight;
      },
      set rowHeight(v: number) {
        writeCalls += 1;
        formatState.rowHeight = v;
      },
      get columnWidth() {
        return formatState.columnWidth;
      },
      autofitColumns: options?.missingAutofit
        ? undefined
        : () => {
            writeCalls += 1;
            formatState.columnWidth = 12;
          },
      autofitRows: options?.missingAutofit
        ? undefined
        : () => {
            writeCalls += 1;
            formatState.rowHeight = kind === "header" ? formatState.rowHeight : 18;
          },
      load(_props?: string) {
        formatLoaded = true;
      },
    };
    if (options?.missingAutofit) {
      delete (format as { autofitColumns?: unknown }).autofitColumns;
      delete (format as { autofitRows?: unknown }).autofitRows;
    }

    return {
      address:
        sheet.poison?.address !== undefined && kind === "used"
          ? sheet.poison.address
          : `${sheet.name}!${address}`,
      get rowCount() {
        if (sheet.poison?.rowCount !== undefined && kind === "used") return sheet.poison.rowCount as number;
        return sheet.rows;
      },
      get columnCount() {
        if (sheet.poison?.columnCount !== undefined && kind === "used") {
          return sheet.poison.columnCount as number;
        }
        return sheet.cols;
      },
      get text() {
        return sheet.text;
      },
      get values() {
        return [[sheet.text || null]];
      },
      format,
      load(_props?: string) {},
    };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    const sheetObj: Record<string, unknown> = {
      get name() {
        return sheet.name;
      },
      load() {},
      get showGridlines() {
        if (sheet.poison?.showGridlines !== undefined) return sheet.poison.showGridlines as boolean;
        return sheet.showGridlines;
      },
      set showGridlines(v: boolean) {
        writeCalls += 1;
        sheet.showGridlines = v;
      },
      getRange(address: string) {
        const bare = address.replace(/\$/g, "");
        // header row heuristic: same start row range with single row
        const isHeader =
          bare === "A1:B1" ||
          /^[A-Z]+1:[A-Z]+1$/i.test(bare) ||
          bare.toUpperCase().endsWith("1") && bare.includes(":");
        return makeRange(sheet, isHeader ? "header" : "other", bare);
      },
      getUsedRangeOrNullObject: options?.missingUsedRange
        ? undefined
        : () => {
            if (sheet.empty) {
              return {
                isNullObject: true,
                address: "",
                rowCount: 0,
                columnCount: 0,
                text: "",
                values: [[null]],
                format: makeRange(sheet, "used", "A1").format,
                load() {},
              };
            }
            return {
              isNullObject: false,
              ...makeRange(sheet, "used", sheet.usedAddress),
            };
          },
      freezePanes: options?.missingFreeze
        ? undefined
        : {
            unfreeze() {
              writeCalls += 1;
              sheet.freezeRows = 0;
            },
            freezeRows(n: number) {
              writeCalls += 1;
              sheet.freezeRows = n;
            },
            getLocationOrNullObject() {
              const fr =
                sheet.poison?.freezeRows !== undefined
                  ? (sheet.poison.freezeRows as number)
                  : sheet.freezeRows;
              if (!fr) {
                return {
                  isNullObject: true,
                  address: "",
                  rowCount: 0,
                  columnCount: 0,
                  load() {},
                };
              }
              return {
                isNullObject: false,
                address: `${sheet.name}!A1:A${fr}`,
                rowCount: fr,
                columnCount: 0,
                load() {},
              };
            },
          },
      pageLayout: makePageLayoutObject(sheet.pageLayout, {
        hasPaperSize: true,
        hasZoom: true,
        hasDraftMode: true,
        hasPrintOrder: true,
        hasFirstPageNumber: true,
        hasHeaderMargin: true,
        hasFooterMargin: true,
        hasHeadersFooters: true,
        hasDefaultForAllPages: true,
        queue: (layoutSheet, patch) => {
          layoutSheet.pending = { ...(layoutSheet.pending ?? {}), ...patch };
        },
      }),
    };
    if (options?.missingShowGridlines) {
      delete sheetObj.showGridlines;
    }
    attachPageBreaks(sheetObj, sheet.pageBreaks, () => {});
    return sheetObj;
  }

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel?: unknown }).Excel = {
    run: async <T>(fn: (ctx: unknown) => Promise<T>) => {
      excelRunCalls += 1;
      const context = {
        workbook: {
          name: "Book1",
          load() {},
          worksheets: {
            items: [...sheets.keys()].map((n) => makeSheet(n)),
            load() {},
            getItem(name: string) {
              if (!sheets.has(name)) throw new Error(`ItemNotFound: ${name}`);
              return makeSheet(name);
            },
            getActiveWorksheet() {
              return makeSheet(activeName);
            },
          },
        },
        async sync() {
          for (const sheet of sheets.values()) {
            if (sheet.pageLayout.pending) {
              sheet.pageLayout.committed = {
                ...sheet.pageLayout.committed,
                ...sheet.pageLayout.pending,
              };
              sheet.pageLayout.pending = undefined;
            }
            commitPageBreaks(sheet.pageBreaks);
          }
        },
      };
      // refresh items from live map each run
      (context.workbook.worksheets as { items: unknown[] }).items = [...sheets.keys()].map((n) =>
        makeSheet(n),
      );
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
          isSetSupported: () => {
            throw new Error("boom");
          },
        },
      },
    };
  } else {
    (globalThis as unknown as {
      Office: { context: { requirements: { isSetSupported: (n: string, v?: string) => boolean } } };
    }).Office = {
      context: {
        requirements: {
          isSetSupported: (_n: string, v?: string) => {
            if (v === "1.8") return excelApi18;
            if (v === "1.9") return excelApi19;
            return true;
          },
        },
      },
    };
  }

  return {
    excelRunCalls: () => excelRunCalls,
    writeCalls: () => writeCalls,
    getSheet: (name: string) => sheets.get(name),
    setPoison: (name: string, poison: SheetState["poison"]) => {
      const s = sheets.get(name);
      if (s) s.poison = poison;
    },
    setEmpty: (name: string, empty: boolean) => {
      const s = sheets.get(name);
      if (s) s.empty = empty;
    },
  };
}
