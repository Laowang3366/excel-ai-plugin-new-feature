/**
 * Sync-gated fake for workbook.template.apply / capture.
 * PropertyNotLoaded until context.sync; everReadBeforeSync is sticky.
 */
import {
  defaultPageLayoutState,
  type PageLayoutSheetState,
} from "./officeJsPageLayoutFakeLayout";
import {
  commitPageBreaks,
  defaultPageBreakSheetState,
  type PageBreakSheetState,
} from "./officeJsPageBreaksFake";

type FormatState = {
  font: { name: any; size: any; bold: any; color: any };
  fill: { color: any };
  horizontalAlignment: any;
  wrapText: any;
  rowHeight: any;
  columnWidth: number;
};

type Poison = Partial<{
  fontName: unknown;
  fontSize: unknown;
  headerFontName: unknown;
  headerFontSize: unknown;
  headerFill: unknown;
  headerFontColor: unknown;
  headerBold: unknown;
  headerAlignment: unknown;
  headerWrap: unknown;
  headerRowHeight: unknown;
  showGridlines: unknown;
  freezeRows: unknown;
  freezeAddress: unknown;
  freezeColumnCount: unknown;
  address: unknown;
  rowCount: unknown;
  columnCount: unknown;
  isNullObject: unknown;
  baseFontName: unknown;
  baseFontSize: unknown;
  baseFontColor: unknown;
  printOrientation: unknown;
  printPaperSize: unknown;
  printFitWide: unknown;
  printFitTall: unknown;
  printFitWideMissing: true;
  printFitTallMissing: true;
  printZoom: unknown; // null | undefined | object override
  printAreaIsNull: unknown;
  printAreaAddress: unknown;
  printHeader: unknown;
  printFooter: unknown;
}>;

type SheetState = {
  name: string;
  empty: boolean;
  showGridlines: any;
  freezeRows: number;
  usedAddress: string;
  rows: number;
  cols: number;
  text: string;
  usedFormat: FormatState;
  headerFormat: FormatState;
  pageLayout: PageLayoutSheetState;
  pageBreaks: PageBreakSheetState;
  isNullObjectOverride?: unknown;
  poison?: Poison;
  postWritePoison?: Partial<{ address: unknown; rowCount: unknown; columnCount: unknown }>;
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

function notLoaded(field: string): never {
  throw new Error(`PropertyNotLoaded: ${field}`);
}

export function installTemplateExcel(options?: {
  excelApi18?: boolean;
  excelApi19?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  missingShowGridlines?: boolean;
  missingFreeze?: boolean;
  missingGetLocation?: boolean;
  missingAutofit?: boolean;
  missingUsedRange?: boolean;
  missingGetRange?: boolean;
  missingRangeLoad?: boolean;
  missingFormat?: boolean;
  missingFormatLoad?: boolean;
  missingFont?: boolean;
  missingFontLoad?: boolean;
  missingFill?: boolean;
  missingFillLoad?: boolean;
  missingDefaultHeaderFooterLoad?: boolean;
  sheetCount?: number;
  activeSheet?: string;
  extraSheets?: Array<{
    name: string;
    empty?: boolean;
    usedAddress?: string;
    rows?: number;
    cols?: number;
    text?: string;
  }>;
}) {
  const excelApi18 = options?.excelApi18 !== false;
  const excelApi19 = options?.excelApi19 !== false;
  let excelRunCalls = 0;
  let writeCalls = 0;
  let syncCount = 0;
  let bulkTextReadCalls = 0;
  let bulkValuesReadCalls = 0;
  let everReadBeforeSync = false;
  let generation = 0;

  const sheets = new Map<string, SheetState>();
  function seed(
    name: string,
    empty = false,
    usedAddress = empty ? "A1" : "A1:B2",
    rows = empty ? 1 : 2,
    cols = empty ? 1 : 2,
    text = empty ? "" : "x",
  ) {
    sheets.set(name, {
      name,
      empty,
      showGridlines: true,
      freezeRows: 0,
      usedAddress,
      rows,
      cols,
      text,
      usedFormat: defaultFormat(),
      headerFormat: defaultFormat(),
      pageLayout: { name, committed: defaultPageLayoutState(), pending: undefined },
      pageBreaks: defaultPageBreakSheetState(),
    });
  }

  const count = options?.sheetCount ?? 2;
  if (count > 500) {
    for (let i = 1; i <= count; i += 1) seed(`S${i}`);
  } else if (count === 1) {
    seed("Sheet1", false);
  } else if (count === 50) {
    for (let i = 1; i <= 50; i += 1) seed(`S${i}`);
  } else {
    seed("Sheet1", false);
    seed("Sheet2", false);
    seed("Empty", true);
  }
  for (const extra of options?.extraSheets ?? []) {
    seed(
      extra.name,
      extra.empty === true,
      extra.usedAddress,
      extra.rows,
      extra.cols,
      extra.text,
    );
  }

  const activeName = options?.activeSheet ?? [...sheets.keys()][0]!;

  type Tracked = {
    queued: Set<string>;
    loaded: Set<string>;
    markLoad: (props: string) => void;
    ensure: (key: string, field: string) => void;
  };

  function makeTracked(): Tracked {
    const queued = new Set<string>();
    const loaded = new Set<string>();
    return {
      queued,
      loaded,
      markLoad(props: string) {
        for (const p of props.split(",").map((s) => s.trim()).filter(Boolean)) {
          queued.add(p);
        }
      },
      ensure(key: string, field: string) {
        if (!loaded.has(key)) {
          everReadBeforeSync = true;
          notLoaded(field);
        }
      },
    };
  }

  const pendingTrackers: Tracked[] = [];

  function commitLoads() {
    for (const t of pendingTrackers) {
      for (const k of t.queued) t.loaded.add(k);
      t.queued.clear();
    }
  }

  function makeFormat(sheet: SheetState, kind: "used" | "header" | "other", needAutofit: boolean) {
    const formatState = kind === "header" ? sheet.headerFormat : sheet.usedFormat;
    const track = makeTracked();
    pendingTrackers.push(track);
    const fontTrack = makeTracked();
    pendingTrackers.push(fontTrack);
    const fillTrack = makeTracked();
    pendingTrackers.push(fillTrack);

    const font: any = {
      load(props?: string) {
        fontTrack.markLoad(props ?? "name,size,bold,color");
      },
      get name() {
        fontTrack.ensure("name", "font.name");
        if (kind === "used" && sheet.poison?.fontName !== undefined) return sheet.poison.fontName;
        if (kind === "header" && sheet.poison?.headerFontName !== undefined) {
          return sheet.poison.headerFontName;
        }
        if (kind === "used" && sheet.poison?.baseFontName !== undefined) {
          return sheet.poison.baseFontName;
        }
        return formatState.font.name;
      },
      set name(v: any) {
        writeCalls += 1;
        formatState.font.name = v as any;
      },
      get size() {
        fontTrack.ensure("size", "font.size");
        if (kind === "used" && sheet.poison?.fontSize !== undefined) return sheet.poison.fontSize;
        if (kind === "header" && sheet.poison?.headerFontSize !== undefined) {
          return sheet.poison.headerFontSize;
        }
        if (kind === "used" && sheet.poison?.baseFontSize !== undefined) {
          return sheet.poison.baseFontSize;
        }
        return formatState.font.size;
      },
      set size(v: any) {
        writeCalls += 1;
        formatState.font.size = v as any;
      },
      get bold() {
        fontTrack.ensure("bold", "font.bold");
        if (kind === "header" && sheet.poison?.headerBold !== undefined) {
          return sheet.poison.headerBold;
        }
        return formatState.font.bold;
      },
      set bold(v: any) {
        writeCalls += 1;
        formatState.font.bold = v as any;
      },
      get color() {
        fontTrack.ensure("color", "font.color");
        if (kind === "header" && sheet.poison?.headerFontColor !== undefined) {
          return sheet.poison.headerFontColor;
        }
        if (kind === "used" && sheet.poison?.baseFontColor !== undefined) {
          return sheet.poison.baseFontColor;
        }
        return formatState.font.color;
      },
      set color(v: any) {
        writeCalls += 1;
        formatState.font.color = v as any;
      },
    };
    if (options?.missingFontLoad) {
      delete font.load;
    }

    const fill: any = {
      load(props?: string) {
        fillTrack.markLoad(props ?? "color");
      },
      get color() {
        fillTrack.ensure("color", "fill.color");
        if (kind === "header" && sheet.poison?.headerFill !== undefined) {
          return sheet.poison.headerFill;
        }
        return formatState.fill.color;
      },
      set color(v: any) {
        writeCalls += 1;
        formatState.fill.color = v as any;
      },
    };
    if (options?.missingFillLoad) {
      delete fill.load;
    }

    if (options?.missingFont) {
      // omit font
    }
    if (options?.missingFill) {
      // omit fill
    }

    const format: any = {
      get font() {
        if (options?.missingFont) return undefined;
        return font;
      },
      get fill() {
        if (options?.missingFill) return undefined;
        return fill;
      },
      get horizontalAlignment() {
        track.ensure("horizontalAlignment", "format.horizontalAlignment");
        if (kind === "header" && sheet.poison?.headerAlignment !== undefined) {
          return sheet.poison.headerAlignment;
        }
        return formatState.horizontalAlignment;
      },
      set horizontalAlignment(v: any) {
        writeCalls += 1;
        formatState.horizontalAlignment = v as any;
      },
      get wrapText() {
        track.ensure("wrapText", "format.wrapText");
        if (kind === "header" && sheet.poison?.headerWrap !== undefined) {
          return sheet.poison.headerWrap;
        }
        return formatState.wrapText;
      },
      set wrapText(v: any) {
        writeCalls += 1;
        formatState.wrapText = v as any;
      },
      get rowHeight() {
        track.ensure("rowHeight", "format.rowHeight");
        if (kind === "header" && sheet.poison?.headerRowHeight !== undefined) {
          return sheet.poison.headerRowHeight;
        }
        return formatState.rowHeight;
      },
      set rowHeight(v: any) {
        writeCalls += 1;
        formatState.rowHeight = v as any;
      },
      get columnWidth() {
        return formatState.columnWidth;
      },
      load(props?: string) {
        track.markLoad(props ?? "horizontalAlignment,wrapText,rowHeight");
      },
    };
    if (options?.missingFormatLoad) {
      delete format.load;
    }

    if (!options?.missingAutofit && needAutofit !== false) {
      format.autofitColumns = () => {
        writeCalls += 1;
        formatState.columnWidth = 12;
      };
      format.autofitRows = () => {
        writeCalls += 1;
      };
    }
    if (options?.missingAutofit) {
      // no autofit methods
    }
    if (options?.missingFormat) return undefined as unknown as Record<string, unknown>;
    return format;
  }

  function makeRange(sheet: SheetState, kind: "used" | "header" | "other", address: string) {
    const dims = makeTracked();
    pendingTrackers.push(dims);
    const textTrack = makeTracked();
    pendingTrackers.push(textTrack);
    const valuesTrack = makeTracked();
    pendingTrackers.push(valuesTrack);
    const format = makeFormat(sheet, kind, kind === "used");

    const qualified =
      sheet.name.includes("!") || /[\s']/.test(sheet.name)
        ? `'${sheet.name.replace(/'/g, "''")}'!${address}`
        : `${sheet.name}!${address}`;

    const range: any = {
      get address() {
        dims.ensure("address", "range.address");
        if (kind === "used" && writeCalls > 0 && sheet.postWritePoison?.address !== undefined) {
          return sheet.postWritePoison.address;
        }
        if (kind === "used" && sheet.poison?.address !== undefined) return sheet.poison.address;
        return qualified;
      },
      get rowCount() {
        dims.ensure("rowCount", "range.rowCount");
        if (kind === "used" && writeCalls > 0 && sheet.postWritePoison?.rowCount !== undefined) {
          return sheet.postWritePoison.rowCount;
        }
        if (kind === "used" && sheet.poison?.rowCount !== undefined) return sheet.poison.rowCount;
        return sheet.rows;
      },
      get columnCount() {
        dims.ensure("columnCount", "range.columnCount");
        if (
          kind === "used" &&
          writeCalls > 0 &&
          sheet.postWritePoison?.columnCount !== undefined
        ) {
          return sheet.postWritePoison.columnCount;
        }
        if (kind === "used" && sheet.poison?.columnCount !== undefined) {
          return sheet.poison.columnCount;
        }
        return sheet.cols;
      },
      get text() {
        textTrack.ensure("text", "range.text");
        if (kind === "used") {
          bulkTextReadCalls += 1;
          if (sheet.rows > 1 || sheet.cols > 1) {
            // multi-cell text still counts as bulk
          }
        }
        return sheet.text;
      },
      get values() {
        valuesTrack.ensure("values", "range.values");
        if (kind === "used") bulkValuesReadCalls += 1;
        return [[sheet.text || null]];
      },
      get format() {
        return format;
      },
      load(props?: string) {
        const p = props ?? "";
        dims.markLoad(p);
        if (p.includes("text")) textTrack.markLoad("text");
        if (p.includes("values")) valuesTrack.markLoad("values");
        if (p.includes("isNullObject")) dims.markLoad("isNullObject");
        if (!p) {
          dims.markLoad("address,rowCount,columnCount,isNullObject");
        }
      },
    };
    if (options?.missingRangeLoad) {
      delete range.load;
    }
    return range;
  }

  function makeUsedNullObject(sheet: SheetState) {
    const dims = makeTracked();
    pendingTrackers.push(dims);
    const format = makeFormat(sheet, "used", true);
    const isNull =
      sheet.isNullObjectOverride !== undefined
        ? sheet.isNullObjectOverride
        : sheet.poison?.isNullObject !== undefined
          ? sheet.poison.isNullObject
          : sheet.empty;

    const obj: any = {
      get isNullObject() {
        // Office.js nullObject: isNullObject is readable without scalar load.
        return isNull;
      },
      get address() {
        dims.ensure("address", "UsedRange.address");
        if (writeCalls > 0 && sheet.postWritePoison?.address !== undefined) {
          return sheet.postWritePoison.address;
        }
        if (sheet.poison?.address !== undefined) return sheet.poison.address;
        const address = sheet.usedAddress;
        return sheet.name.includes("!") || /[\s']/.test(sheet.name)
          ? `'${sheet.name.replace(/'/g, "''")}'!${address}`
          : `${sheet.name}!${address}`;
      },
      get rowCount() {
        dims.ensure("rowCount", "UsedRange.rowCount");
        if (writeCalls > 0 && sheet.postWritePoison?.rowCount !== undefined) {
          return sheet.postWritePoison.rowCount;
        }
        if (sheet.poison?.rowCount !== undefined) return sheet.poison.rowCount;
        return sheet.rows;
      },
      get columnCount() {
        dims.ensure("columnCount", "UsedRange.columnCount");
        if (writeCalls > 0 && sheet.postWritePoison?.columnCount !== undefined) {
          return sheet.postWritePoison.columnCount;
        }
        if (sheet.poison?.columnCount !== undefined) return sheet.poison.columnCount;
        return sheet.cols;
      },
      get text() {
        dims.ensure("text", "UsedRange.text");
        bulkTextReadCalls += 1;
        return sheet.text;
      },
      get values() {
        dims.ensure("values", "UsedRange.values");
        bulkValuesReadCalls += 1;
        return [[sheet.text || null]];
      },
      get format() {
        return format;
      },
      load(props?: string) {
        const p = props ?? "isNullObject,address,rowCount,columnCount";
        dims.markLoad(p);
      },
    };
    if (options?.missingRangeLoad) {
      delete obj.load;
    }
    return obj;
  }

  function makeFreezeLoc(sheet: SheetState) {
    const dims = makeTracked();
    pendingTrackers.push(dims);
    const fr =
      sheet.poison?.freezeRows !== undefined
        ? (sheet.poison.freezeRows as number)
        : sheet.freezeRows;
    return {
      get isNullObject() {
        dims.ensure("isNullObject", "freeze.isNullObject");
        return !fr;
      },
      get address() {
        dims.ensure("address", "freeze.address");
        if (sheet.poison?.freezeAddress !== undefined) return sheet.poison.freezeAddress;
        return `${sheet.name}!A1:A${fr || 1}`;
      },
      get rowCount() {
        dims.ensure("rowCount", "freeze.rowCount");
        return fr || 0;
      },
      get columnCount() {
        dims.ensure("columnCount", "freeze.columnCount");
        if (sheet.poison?.freezeColumnCount !== undefined) return sheet.poison.freezeColumnCount;
        return 0;
      },
      load(props?: string) {
        dims.markLoad(props ?? "isNullObject,address,rowCount,columnCount");
      },
    };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    const nameTrack = makeTracked();
    pendingTrackers.push(nameTrack);
    const gridTrack = makeTracked();
    pendingTrackers.push(gridTrack);

    const sheetObj: Record<string, unknown> = {
      get name() {
        // name often loaded via items/name or sheet.load("name")
        if (!nameTrack.loaded.has("name")) {
          // collection load items/name commits name for sheet items
          if (!nameTrack.loaded.has("name")) {
            // allow if collection seeded — mark on collection load
          }
        }
        if (!nameTrack.loaded.has("name")) {
          everReadBeforeSync = true;
          notLoaded("Worksheet.name");
        }
        return sheet.name;
      },
      load(props?: string) {
        nameTrack.markLoad(props ?? "name");
        if (props?.includes("showGridlines")) gridTrack.markLoad("showGridlines");
        if (!props || props.includes("name")) nameTrack.markLoad("name");
      },
      get showGridlines() {
        gridTrack.ensure("showGridlines", "showGridlines");
        if (sheet.poison?.showGridlines !== undefined) return sheet.poison.showGridlines;
        return sheet.showGridlines;
      },
      set showGridlines(v: any) {
        writeCalls += 1;
        sheet.showGridlines = Boolean(v);
      },
      getRange: options?.missingGetRange
        ? undefined
        : (address: string) => {
            const bare = address.replace(/\$/g, "");
            const isHeader =
              bare === "A1:B1" ||
              /^[A-Z]+1:[A-Z]+1$/i.test(bare) ||
              (bare.toUpperCase().endsWith("1") && bare.includes(":"));
            return makeRange(sheet, isHeader ? "header" : "other", bare);
          },
      getUsedRangeOrNullObject: options?.missingUsedRange
        ? undefined
        : () => makeUsedNullObject(sheet),
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
            getLocationOrNullObject: options?.missingGetLocation
              ? undefined
              : () => makeFreezeLoc(sheet),
          },
      get pageLayout() {
        // Capture print subset only (no full pageLayout / page breaks).
        const layoutTrack = makeTracked();
        pendingTrackers.push(layoutTrack);
        const areaTrack = makeTracked();
        pendingTrackers.push(areaTrack);
        const rowsTrack = makeTracked();
        pendingTrackers.push(rowsTrack);
        const colsTrack = makeTracked();
        pendingTrackers.push(colsTrack);
        const hfTrack = makeTracked();
        pendingTrackers.push(hfTrack);

        const printArea = () => ({
          get isNullObject() {
            areaTrack.ensure("isNullObject", "printArea.isNullObject");
            if (sheet.poison?.printAreaIsNull !== undefined) return sheet.poison.printAreaIsNull;
            return true;
          },
          get address() {
            areaTrack.ensure("address", "printArea.address");
            if (sheet.poison?.printAreaAddress !== undefined) return sheet.poison.printAreaAddress;
            return `${sheet.name}!$A$1:$B$2`;
          },
          load(props?: string) {
            areaTrack.markLoad(props ?? "isNullObject,address");
          },
        });
        const titleRows = () => ({
          get isNullObject() {
            rowsTrack.ensure("isNullObject", "titleRows.isNullObject");
            return true;
          },
          get address() {
            rowsTrack.ensure("address", "titleRows.address");
            return `${sheet.name}!$1:$1`;
          },
          load(props?: string) {
            rowsTrack.markLoad(props ?? "isNullObject,address");
          },
        });
        const titleCols = () => ({
          get isNullObject() {
            colsTrack.ensure("isNullObject", "titleCols.isNullObject");
            return true;
          },
          get address() {
            colsTrack.ensure("address", "titleCols.address");
            return `${sheet.name}!$A:$A`;
          },
          load(props?: string) {
            colsTrack.markLoad(props ?? "isNullObject,address");
          },
        });

        const def: any = {
          get centerHeader() {
            hfTrack.ensure("centerHeader", "centerHeader");
            if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printHeader")
            ) {
              return sheet.poison.printHeader;
            }
            return "";
          },
          get centerFooter() {
            hfTrack.ensure("centerFooter", "centerFooter");
            if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printFooter")
            ) {
              return sheet.poison.printFooter;
            }
            return "";
          },
          load(props?: string) {
            hfTrack.markLoad(props ?? "centerHeader,centerFooter");
          },
        };
        if (options?.missingDefaultHeaderFooterLoad) {
          delete def.load;
        }

        return {
          load(props?: string) {
            layoutTrack.markLoad(props ?? "orientation,paperSize,zoom");
          },
          get orientation() {
            layoutTrack.ensure("orientation", "layout.orientation");
            if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printOrientation")
            ) {
              return sheet.poison.printOrientation;
            }
            return "Portrait";
          },
          get paperSize() {
            layoutTrack.ensure("paperSize", "layout.paperSize");
            if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printPaperSize")
            ) {
              return sheet.poison.printPaperSize;
            }
            return "A4";
          },
          get zoom() {
            layoutTrack.ensure("zoom", "layout.zoom");
            if (sheet.poison && Object.prototype.hasOwnProperty.call(sheet.poison, "printZoom")) {
              return sheet.poison.printZoom;
            }
            const zoomObj: Record<string, unknown> = {};
            if (sheet.poison?.printFitWideMissing) {
              // omit horizontalFitToPages key
            } else if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printFitWide")
            ) {
              zoomObj.horizontalFitToPages = sheet.poison.printFitWide;
            } else {
              zoomObj.horizontalFitToPages = null;
            }
            if (sheet.poison?.printFitTallMissing) {
              // omit verticalFitToPages key
            } else if (
              sheet.poison &&
              Object.prototype.hasOwnProperty.call(sheet.poison, "printFitTall")
            ) {
              zoomObj.verticalFitToPages = sheet.poison.printFitTall;
            } else {
              zoomObj.verticalFitToPages = null;
            }
            return zoomObj;
          },
          get headersFooters() {
            return { defaultForAllPages: def };
          },
          getPrintAreaOrNullObject: printArea,
          getPrintTitleRowsOrNullObject: titleRows,
          getPrintTitleColumnsOrNullObject: titleCols,
        };
      },
    };

    if (options?.missingShowGridlines) {
      delete sheetObj.showGridlines;
    }

    // seed name as loadable via collection
    (sheetObj as { __nameTrack: Tracked }).__nameTrack = nameTrack;
    return sheetObj;
  }

  const sheetObjects = new Map<string, ReturnType<typeof makeSheet>>();
  for (const name of sheets.keys()) {
    sheetObjects.set(name, makeSheet(name));
  }
  // Allow reading names after items/name load
  function markAllNamesQueued() {
    for (const obj of sheetObjects.values()) {
      const t = (obj as { __nameTrack?: Tracked }).__nameTrack;
      t?.markLoad("name");
    }
  }

  const context = {
    workbook: {
      get name() {
        wbNameTrack.ensure("name", "Workbook.name");
        return "Book1";
      },
      load(props?: string) {
        if (props?.includes("name")) wbNameTrack.markLoad("name");
      },
      worksheets: {
        load(props?: string) {
          if (props?.includes("items") || props?.includes("name")) {
            itemsTrack.markLoad("items");
            markAllNamesQueued();
          }
        },
        get items() {
          itemsTrack.ensure("items", "worksheets.items");
          return [...sheetObjects.values()];
        },
        getActiveWorksheet() {
          return sheetObjects.get(activeName)!;
        },
        getItem(name: string) {
          const s = sheetObjects.get(name);
          if (!s) throw new Error(`sheet not found: ${name}`);
          return s;
        },
      },
    },
    async sync() {
      syncCount += 1;
      generation += 1;
      commitLoads();
      for (const st of sheets.values()) {
        commitPageBreaks(st.pageBreaks);
        if (st.pageLayout.pending) {
          st.pageLayout.committed = { ...st.pageLayout.committed, ...st.pageLayout.pending };
          st.pageLayout.pending = undefined;
        }
      }
    },
  };

  const itemsTrack = makeTracked();
  pendingTrackers.push(itemsTrack);
  const wbNameTrack = makeTracked();
  pendingTrackers.push(wbNameTrack);

  // Node vitest: getExcelRun() reads window.Excel only — mirror baseline install.
  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: (fn: (ctx: typeof context) => Promise<unknown>) => Promise<unknown> } }).Excel = {
    run: async (fn) => {
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
          isSetSupported: () => {
            throw new Error("boom");
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
    syncCount: () => syncCount,
    bulkTextReadCalls: () => bulkTextReadCalls,
    bulkValuesReadCalls: () => bulkValuesReadCalls,
    everReadBeforeSync: () => everReadBeforeSync,
    generation: () => generation,
    getSheet: (name: string) => sheets.get(name),
    setPoison: (name: string, poison: Poison) => {
      const s = sheets.get(name);
      if (s) s.poison = { ...(s.poison ?? {}), ...poison };
    },
    setPostWritePoison: (
      name: string,
      poison: NonNullable<SheetState["postWritePoison"]>,
    ) => {
      const s = sheets.get(name);
      if (s) s.postWritePoison = { ...(s.postWritePoison ?? {}), ...poison };
    },
    setIsNullObject: (name: string, value: unknown) => {
      const s = sheets.get(name);
      if (s) s.isNullObjectOverride = value;
    },
    setEmpty: (name: string, empty: boolean) => {
      const s = sheets.get(name);
      if (s) {
        s.empty = empty;
        if (empty) {
          s.rows = 1;
          s.cols = 1;
          s.text = "";
          s.usedAddress = "A1";
        }
      }
    },
    setUsedDims: (name: string, rows: number, cols: number, address?: string, text?: string) => {
      const s = sheets.get(name);
      if (!s) return;
      s.rows = rows;
      s.cols = cols;
      if (address) s.usedAddress = address;
      if (text !== undefined) s.text = text;
      s.empty = false;
    },
    setBaseNull: (name: string) => {
      const s = sheets.get(name);
      if (!s) return;
      s.usedFormat.font.name = null;
      s.usedFormat.font.size = null;
      s.usedFormat.font.color = null;
    },
    /** Force a premature property read for gate self-test. */
    forcePrematureRead: (_name: string) => {
      const used = (sheetObjects.get(_name) as { getUsedRangeOrNullObject?: () => { address: unknown } })
        ?.getUsedRangeOrNullObject?.();
      try {
        void used?.address;
      } catch {
        // expected
      }
    },
  };
}
