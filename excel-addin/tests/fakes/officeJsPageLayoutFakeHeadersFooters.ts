/** Sync-gated headersFooters.defaultForAllPages for pageLayout fake. */
import type { PageLayoutSheetState, PageLayoutState } from "./officeJsPageLayoutFakeLayout";

export type HeadersFootersState = {
  leftHeader: string;
  centerHeader: string;
  rightHeader: string;
  leftFooter: string;
  centerFooter: string;
  rightFooter: string;
};

export type MakeHeadersFootersOptions = {
  hasHeadersFooters: boolean;
  hasDefaultForAllPages: boolean;
  /** When set, omit that slot property from defaultForAllPages. */
  missingSlot?: keyof HeadersFootersState;
  queue: (sheet: PageLayoutSheetState, patch: Partial<PageLayoutState>) => void;
};

function slotProp(
  sheet: PageLayoutSheetState,
  key: keyof HeadersFootersState,
  queue: MakeHeadersFootersOptions["queue"],
) {
  return {
    enumerable: true,
    configurable: true,
    get() {
      return sheet.committed[key];
    },
    set(v: string) {
      queue(sheet, { [key]: v } as Partial<PageLayoutState>);
    },
  };
}

export function attachHeadersFooters(
  layout: Record<string, unknown>,
  sheet: PageLayoutSheetState,
  options: MakeHeadersFootersOptions,
): void {
  if (!options.hasHeadersFooters) return;

  const defaultPage: Record<string, unknown> = {
    load() {},
  };
  const slots: Array<keyof HeadersFootersState> = [
    "leftHeader",
    "centerHeader",
    "rightHeader",
    "leftFooter",
    "centerFooter",
    "rightFooter",
  ];
  for (const slot of slots) {
    if (options.missingSlot === slot) continue;
    Object.defineProperty(defaultPage, slot, slotProp(sheet, slot, options.queue));
  }

  const group: Record<string, unknown> = {
    load() {},
  };
  if (options.hasDefaultForAllPages) {
    Object.defineProperty(group, "defaultForAllPages", {
      enumerable: true,
      configurable: true,
      get() {
        return defaultPage;
      },
    });
  }

  Object.defineProperty(layout, "headersFooters", {
    enumerable: true,
    configurable: true,
    get() {
      return group;
    },
  });
}
