/** Sync-gated Worksheet horizontal/vertical page break collections. */
export type PageBreakCell = { address: string };

export type PageBreakSheetState = {
  horizontal: PageBreakCell[];
  vertical: PageBreakCell[];
  pendingHorizontal?: PageBreakCell[] | "clear" | undefined;
  pendingVertical?: PageBreakCell[] | "clear" | undefined;
  /** When false, omit collection member. */
  hasHorizontal?: boolean;
  hasVertical?: boolean;
  hasAdd?: boolean;
  hasRemove?: boolean;
  hasItems?: boolean;
  hasGetCellAfterBreak?: boolean;
};

export function defaultPageBreakSheetState(): PageBreakSheetState {
  return {
    horizontal: [],
    vertical: [],
    hasHorizontal: true,
    hasVertical: true,
    hasAdd: true,
    hasRemove: true,
    hasItems: true,
    hasGetCellAfterBreak: true,
  };
}

function makeCollection(
  state: PageBreakSheetState,
  direction: "horizontal" | "vertical",
  onWrite: () => void,
) {
  const collection: Record<string, unknown> = {
    load() {},
  };

  if (state.hasItems !== false) {
    Object.defineProperty(collection, "items", {
      enumerable: true,
      configurable: true,
      get() {
        return (direction === "horizontal" ? state.horizontal : state.vertical).map((cell) => {
          const item: Record<string, unknown> = {};
          if (state.hasGetCellAfterBreak !== false) {
            item.getCellAfterBreak = () => ({
              address: cell.address,
              load() {},
            });
          }
          return item;
        });
      },
    });
  }

  if (state.hasAdd !== false) {
    collection.add = (reference: string) => {
      onWrite();
      const cell = { address: String(reference) };
      if (direction === "horizontal") {
        const base =
          state.pendingHorizontal === "clear"
            ? []
            : (state.pendingHorizontal ?? state.horizontal);
        state.pendingHorizontal = [...base, cell];
      } else {
        const base =
          state.pendingVertical === "clear" ? [] : (state.pendingVertical ?? state.vertical);
        state.pendingVertical = [...base, cell];
      }
      return {
        getCellAfterBreak() {
          return { address: cell.address, load() {} };
        },
      };
    };
  }

  if (state.hasRemove !== false) {
    collection.removePageBreaks = () => {
      onWrite();
      if (direction === "horizontal") state.pendingHorizontal = "clear";
      else state.pendingVertical = "clear";
    };
  }

  return collection;
}

export function attachPageBreaks(
  sheetObj: Record<string, unknown>,
  state: PageBreakSheetState,
  onWrite: () => void,
): void {
  if (state.hasHorizontal !== false) {
    Object.defineProperty(sheetObj, "horizontalPageBreaks", {
      enumerable: true,
      configurable: true,
      get() {
        return makeCollection(state, "horizontal", onWrite);
      },
    });
  }
  if (state.hasVertical !== false) {
    Object.defineProperty(sheetObj, "verticalPageBreaks", {
      enumerable: true,
      configurable: true,
      get() {
        return makeCollection(state, "vertical", onWrite);
      },
    });
  }
}

export function commitPageBreaks(state: PageBreakSheetState): void {
  if (state.pendingHorizontal !== undefined) {
    state.horizontal =
      state.pendingHorizontal === "clear" ? [] : state.pendingHorizontal;
    state.pendingHorizontal = undefined;
  }
  if (state.pendingVertical !== undefined) {
    state.vertical = state.pendingVertical === "clear" ? [] : state.pendingVertical;
    state.pendingVertical = undefined;
  }
}
