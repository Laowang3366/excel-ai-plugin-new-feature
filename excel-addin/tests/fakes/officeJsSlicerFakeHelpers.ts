/** Proxy helpers for Office.js slicer fake (PropertyNotLoaded until load+sync). */

export type SlicerFakeItem = {
  key: string;
  name: string;
  isSelected: boolean;
  hasData: boolean;
};

export type SlicerFakeState = {
  id: string;
  name: string;
  caption: string;
  sheetName: string;
  top: number;
  left: number;
  width: number;
  height: number;
  sortBy: string;
  style: string;
  isFilterCleared: boolean;
  items: SlicerFakeItem[];
  /** Optional type-poison overrides applied after load (test-only). */
  poison?: Partial<Record<string, unknown>>;
  itemPoison?: Array<Partial<Record<string, unknown>> | undefined>;
};

export type WriteCounts = {
  add: number;
  delete: number;
  selectItems: number;
  clearFilters: number;
  propertySets: number;
};

export type ClientResultBag = {
  pending: Array<{ resolve: () => void }>;
  generation: number;
};

export function notLoaded(name: string): Error {
  const err = new Error(`PropertyNotLoaded: ${name}`);
  (err as { code?: string }).code = "PropertyNotLoaded";
  return err;
}

export type LoadMaps = {
  pending: WeakMap<object, Set<string>>;
  loaded: WeakMap<object, Set<string>>;
  proxies: object[];
};

export function createLoadMaps(): LoadMaps {
  return { pending: new WeakMap(), loaded: new WeakMap(), proxies: [] };
}

export function markLoad(maps: LoadMaps, target: object, props: string): void {
  const set = maps.pending.get(target) ?? new Set<string>();
  for (const p of props.split(",").map((x) => x.trim()).filter(Boolean)) set.add(p);
  maps.pending.set(target, set);
  if (!maps.proxies.includes(target)) maps.proxies.push(target);
}

export function commitLoads(maps: LoadMaps): void {
  for (const proxy of maps.proxies) {
    const pending = maps.pending.get(proxy);
    if (!pending) continue;
    const done = maps.loaded.get(proxy) ?? new Set<string>();
    for (const p of pending) done.add(p);
    maps.loaded.set(proxy, done);
    maps.pending.set(proxy, new Set());
  }
}

export function requireLoaded(maps: LoadMaps, target: object, prop: string): void {
  const done = maps.loaded.get(target);
  if (!done || !done.has(prop)) throw notLoaded(prop);
}

export function makeClientResult(getValue: () => string[], bag: ClientResultBag): { value: string[] } {
  let ready = false;
  bag.pending.push({
    resolve: () => {
      ready = true;
    },
  });
  return {
    get value() {
      if (!ready) throw notLoaded("ClientResult.value");
      return getValue();
    },
  };
}

export function flushClientResults(bag: ClientResultBag): void {
  const pending = bag.pending.splice(0, bag.pending.length);
  for (const p of pending) p.resolve();
}

export type SlicerProxyOpts = {
  poisonSortBy?: Map<string, string>;
  noGetSelected?: Set<string>;
  selectItemsNoOp?: Set<string>;
  clientResults: ClientResultBag;
  onDelete?: () => void;
};

export function makeSlicerProxy(
  s: SlicerFakeState,
  maps: LoadMaps,
  writeCounts: WriteCounts,
  opts: SlicerProxyOpts,
): Record<string, unknown> {
  const proxy: Record<string, unknown> = {
    load(props: string) {
      markLoad(maps, proxy, props);
    },
    delete() {
      writeCounts.delete += 1;
      opts.onDelete?.();
    },
    clearFilters() {
      writeCounts.clearFilters += 1;
      for (const it of s.items) it.isSelected = true;
      s.isFilterCleared = true;
    },
    selectItems(items?: string[]) {
      writeCounts.selectItems += 1;
      if (opts.selectItemsNoOp?.has(s.name)) return;
      if (!items || items.length === 0) {
        for (const it of s.items) it.isSelected = true;
        s.isFilterCleared = true;
        return;
      }
      const set = new Set(items);
      for (const it of s.items) it.isSelected = set.has(it.key);
      s.isFilterCleared = false;
    },
  };
  maps.proxies.push(proxy);

  if (!opts.noGetSelected?.has(s.name)) {
    proxy.getSelectedItems = () =>
      makeClientResult(
        () => s.items.filter((i) => i.isSelected).map((i) => i.key),
        opts.clientResults,
      );
  }

  const bind = (prop: string, get: () => unknown, set?: (v: unknown) => void): void => {
    Object.defineProperty(proxy, prop, {
      get() {
        requireLoaded(maps, proxy, prop);
        if (s.poison && Object.prototype.hasOwnProperty.call(s.poison, prop)) {
          return s.poison[prop];
        }
        if (prop === "sortBy" && opts.poisonSortBy?.has(s.name)) {
          return opts.poisonSortBy.get(s.name);
        }
        return get();
      },
      set(v: unknown) {
        writeCounts.propertySets += 1;
        set?.(v);
      },
      configurable: true,
    });
  };

  bind("name", () => s.name, (v) => {
    s.name = String(v);
  });
  bind("id", () => s.id);
  bind("caption", () => s.caption, (v) => {
    s.caption = String(v);
  });
  bind("top", () => s.top, (v) => {
    s.top = Number(v);
  });
  bind("left", () => s.left, (v) => {
    s.left = Number(v);
  });
  bind("width", () => s.width, (v) => {
    s.width = Number(v);
  });
  bind("height", () => s.height, (v) => {
    s.height = Number(v);
  });
  bind("sortBy", () => s.sortBy, (v) => {
    s.sortBy = String(v);
  });
  bind("style", () => s.style, (v) => {
    s.style = String(v);
  });
  bind("isFilterCleared", () => s.isFilterCleared, (v) => {
    s.isFilterCleared = Boolean(v);
  });

  const worksheet = {
    load(props: string) {
      markLoad(maps, worksheet, props);
    },
    get name() {
      requireLoaded(maps, worksheet, "name");
      return s.sheetName;
    },
  };
  maps.proxies.push(worksheet);
  Object.defineProperty(proxy, "worksheet", {
    get() {
      return worksheet;
    },
    configurable: true,
  });

  const itemsColl: { items: object[]; load(props: string): void } = {
    items: [],
    load(props: string) {
      markLoad(maps, itemsColl, props);
    },
  };
  maps.proxies.push(itemsColl);
  Object.defineProperty(itemsColl, "items", {
    get() {
      requireLoaded(maps, itemsColl, "items");
      return s.items.map((item, index) => makeItemProxy(item, maps, writeCounts, s, index));
    },
    configurable: true,
  });
  Object.defineProperty(proxy, "slicerItems", {
    get() {
      return itemsColl;
    },
    configurable: true,
  });

  return proxy;
}

function makeItemProxy(
  item: SlicerFakeItem,
  maps: LoadMaps,
  writeCounts: WriteCounts,
  parent: SlicerFakeState,
  index: number,
): Record<string, unknown> {
  const proxy: Record<string, unknown> = {
    load(props: string) {
      markLoad(maps, proxy, props);
    },
  };
  maps.proxies.push(proxy);
  for (const key of ["key", "name", "isSelected", "hasData"] as const) {
    Object.defineProperty(proxy, key, {
      get() {
        requireLoaded(maps, proxy, key);
        const poison = parent.itemPoison?.[index];
        if (poison && Object.prototype.hasOwnProperty.call(poison, key)) {
          return poison[key];
        }
        return item[key];
      },
      set(v: unknown) {
        writeCounts.propertySets += 1;
        (item as Record<string, unknown>)[key] = v;
      },
      configurable: true,
    });
  }
  return proxy;
}

export function defaultItems(): SlicerFakeItem[] {
  return [
    { key: "A", name: "A", isSelected: true, hasData: true },
    { key: "B", name: "B", isSelected: true, hasData: true },
    { key: "C", name: "C", isSelected: true, hasData: true },
  ];
}
