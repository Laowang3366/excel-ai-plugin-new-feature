/** Shared types/helpers for officeJsPivotFake. */
export type HierarchyAxis = "row" | "column" | "filter" | "data";

export type PivotState = {
  name: string;
  sheetName: string;
  sourceAddress: string;
  destAddress: string;
  hierarchyNames: string[];
  row: string[];
  column: string[];
  filter: string[];
  data: Array<{ name: string; summarizeBy: string; caption: string }>;
  refreshed: boolean;
  layoutEndRow: number;
};

export type SheetState = {
  name: string;
  pivots: Map<string, PivotState>;
};

export type Pending =
  | { kind: "add"; sheet: string; pivot: PivotState }
  | {
      kind: "layout";
      pivot: PivotState;
      axis: HierarchyAxis;
      field: string;
      data?: { name: string; summarizeBy: string; caption: string };
    }
  | { kind: "refresh"; pivot: PivotState }
  | { kind: "connectionRefreshAll" }
  | { kind: "createSheet"; name: string };

export function createLoadTracker() {
  const loaded = new WeakMap<object, Set<string>>();
  function markLoaded(obj: object, props: string) {
    const set = loaded.get(obj) ?? new Set<string>();
    for (const p of props.split(",").map((s) => s.trim()).filter(Boolean)) {
      set.add(p);
      if (p.startsWith("items/")) set.add("items");
    }
    loaded.set(obj, set);
  }
  function requireLoaded(obj: object, prop: string, strictLoad: boolean): void {
    if (!strictLoad) return;
    const set = loaded.get(obj);
    if (!set || !set.has(prop)) {
      throw new Error(`PropertyNotLoaded: ${prop}`);
    }
  }
  return { markLoaded, requireLoaded };
}

export function makeRange(
  address: string,
  sheetName: string,
  requireLoaded: (obj: object, prop: string) => void,
  markLoaded: (obj: object, props: string) => void,
) {
  const bare = address.includes("!") ? address.split("!")[1]! : address;
  const full = address.includes("!") ? address : `${sheetName}!${bare}`;
  const range: {
    address: string;
    rowIndex: number;
    rowCount: number;
    load: (props?: string) => void;
  } = {
    get address() {
      requireLoaded(range, "address");
      return full;
    },
    rowIndex: 0,
    rowCount: Math.max(1, Number((/(\d+)$/.exec(bare.split(":")[0] ?? "") ?? [])[1] ?? 1)),
    load(props?: string) {
      if (props) markLoaded(range, props);
      else markLoaded(range, "address,rowIndex,rowCount");
    },
  };
  return range;
}


export type InstallPivotExcelOptions = {
  excelApi18?: boolean;
  excelApi13?: boolean;
  excelApi17?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  hierarchyNames?: string[];
  sheets?: string[];
  tamperHierarchies?: boolean;
  strictLoad?: boolean;
  missingAdd?: boolean;
  missingRefresh?: boolean;
  missingDataConnections?: boolean;
  missingRefreshAll?: boolean;
  noDataSourceString?: boolean;
};

export function installPivotOfficeRequirements(options: {
  excelApi18: boolean;
  excelApi17: boolean;
  excelApi13: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
}): void {
  const g = globalThis as unknown as {
    window: unknown;
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, version?: string) => boolean };
      };
    };
  };
  g.window = globalThis;
  if (options.missingIsSetSupported) {
    g.Office = { context: { requirements: {} } };
    return;
  }
  if (options.isSetSupportedThrows) {
    g.Office = {
      context: {
        requirements: {
          isSetSupported: () => {
            throw new Error("isSetSupported boom");
          },
        },
      },
    };
    return;
  }
  g.Office = {
    context: {
      requirements: {
        isSetSupported: (name: string, version?: string) => {
          if (name !== "ExcelApi") return false;
          if (version === "1.8") return options.excelApi18;
          if (version === "1.7") return options.excelApi17;
          if (version === "1.3") return options.excelApi13;
          return false;
        },
      },
    },
  };
}
