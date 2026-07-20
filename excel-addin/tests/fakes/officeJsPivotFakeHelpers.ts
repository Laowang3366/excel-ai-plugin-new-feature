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
