/** Layout state + pageLayout object factory for sync-gated pageLayout fake. */

export type PageLayoutState = {
  orientation: string;
  centerHorizontally: boolean;
  centerVertically: boolean;
  printGridlines: boolean;
  printHeadings: boolean;
  blackAndWhite: boolean;
  draftMode: boolean;
  printOrder: string;
  /** Host may store number | "" | null for firstPageNumber. */
  firstPageNumber: number | string | null;
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

export type PageLayoutSheetState = {
  /** Host-facing Worksheet.name after load; may be non-string for failure tests. */
  name: unknown;
  committed: PageLayoutState;
  pending: Partial<PageLayoutState> | undefined;
};

export function defaultPageLayoutState(): PageLayoutState {
  return {
    orientation: "Portrait",
    centerHorizontally: false,
    centerVertically: false,
    printGridlines: false,
    printHeadings: false,
    blackAndWhite: false,
    draftMode: false,
    printOrder: "DownThenOver",
    firstPageNumber: null,
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

export type MakePageLayoutOptions = {
  hasPaperSize: boolean;
  hasZoom: boolean;
  hasDraftMode: boolean;
  hasPrintOrder: boolean;
  hasFirstPageNumber: boolean;
  queue: (sheet: PageLayoutSheetState, patch: Partial<PageLayoutState>) => void;
};

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

export function makePageLayoutObject(
  sheet: PageLayoutSheetState,
  options: MakePageLayoutOptions,
): Record<string, unknown> {
  const { queue } = options;
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

  if (options.hasDraftMode) {
    Object.defineProperty(layout, "draftMode", {
      enumerable: true,
      configurable: true,
      get() {
        return sheet.committed.draftMode;
      },
      set(v: boolean) {
        queue(sheet, { draftMode: v });
      },
    });
  }

  if (options.hasPrintOrder) {
    Object.defineProperty(layout, "printOrder", {
      enumerable: true,
      configurable: true,
      get() {
        return sheet.committed.printOrder;
      },
      set(v: string) {
        queue(sheet, { printOrder: v });
      },
    });
  }

  if (options.hasFirstPageNumber) {
    Object.defineProperty(layout, "firstPageNumber", {
      enumerable: true,
      configurable: true,
      get() {
        return sheet.committed.firstPageNumber;
      },
      set(v: number | string | null) {
        queue(sheet, { firstPageNumber: v });
      },
    });
  }

  if (options.hasPaperSize) {
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

  if (options.hasZoom) {
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
      set(zoomOptions: {
        scale?: number | null;
        horizontalFitToPages?: number;
        verticalFitToPages?: number;
      }) {
        if (!zoomOptions || typeof zoomOptions !== "object") return;
        const patch: Partial<PageLayoutState> = {};
        if (Object.prototype.hasOwnProperty.call(zoomOptions, "scale")) {
          patch.zoomScale = zoomOptions.scale ?? null;
          patch.fitToPagesWide = null;
          patch.fitToPagesTall = null;
        }
        if (Object.prototype.hasOwnProperty.call(zoomOptions, "horizontalFitToPages")) {
          patch.fitToPagesWide = zoomOptions.horizontalFitToPages ?? null;
          patch.zoomScale = null;
        }
        if (Object.prototype.hasOwnProperty.call(zoomOptions, "verticalFitToPages")) {
          patch.fitToPagesTall = zoomOptions.verticalFitToPages ?? null;
          patch.zoomScale = null;
        }
        if (Object.keys(patch).length > 0) queue(sheet, patch);
      },
    });
  }

  return layout;
}
