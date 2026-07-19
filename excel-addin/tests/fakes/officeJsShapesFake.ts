/**
 * Sync-gated fake for Worksheet.shapes.
 * Mutations stay pending until context.sync(); reads use committed only.
 */
export function installShapesExcel() {
  type ShapeState = {
    name: string;
    type: string;
    geometricShapeType: string | null;
    left: number;
    top: number;
    width: number;
    height: number;
    visible: boolean;
    hasText: boolean;
    text: string;
  };

  type Entry = {
    committed: ShapeState | null;
    pending: ShapeState | "delete" | undefined;
  };

  type SheetState = {
    name: string;
    byKey: Map<string, Entry>;
    /** Stable order of shape keys (committed names after sync). */
    order: string[];
  };

  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", { name: "Sheet1", byKey: new Map(), order: [] });
  sheets.set("Data", { name: "Data", byKey: new Map(), order: [] });

  let seq = 0;

  function getSheet(name: string): SheetState {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return sheet;
  }

  function clone(state: ShapeState): ShapeState {
    return { ...state };
  }

  function effective(entry: Entry): ShapeState | null {
    if (entry.pending === "delete") return null;
    if (entry.pending !== undefined) return entry.pending;
    return entry.committed;
  }

  function ensurePending(entry: Entry): ShapeState {
    if (entry.pending === "delete") throw new Error("shape pending delete");
    if (entry.pending === undefined) {
      if (!entry.committed) throw new Error("shape has no state");
      entry.pending = clone(entry.committed);
    }
    return entry.pending;
  }

  function makeProxy(sheetName: string, key: string) {
    const sheet = getSheet(sheetName);
    const entry = () => {
      const found = sheet.byKey.get(key);
      if (!found) throw new Error(`missing shape key ${key}`);
      return found;
    };

    return {
      get name() {
        return entry().committed?.name ?? "";
      },
      set name(next: string) {
        const e = entry();
        const p = ensurePending(e);
        p.name = next;
      },
      get type() {
        return entry().committed?.type ?? "GeometricShape";
      },
      get geometricShapeType() {
        return entry().committed?.geometricShapeType ?? null;
      },
      get left() {
        return entry().committed?.left ?? 0;
      },
      set left(v: number) {
        ensurePending(entry()).left = v;
      },
      get top() {
        return entry().committed?.top ?? 0;
      },
      set top(v: number) {
        ensurePending(entry()).top = v;
      },
      get width() {
        return entry().committed?.width ?? 0;
      },
      set width(v: number) {
        ensurePending(entry()).width = v;
      },
      get height() {
        return entry().committed?.height ?? 0;
      },
      set height(v: number) {
        ensurePending(entry()).height = v;
      },
      get visible() {
        return entry().committed?.visible ?? true;
      },
      set visible(v: boolean) {
        ensurePending(entry()).visible = v;
      },
      textFrame: {
        get hasText() {
          return entry().committed?.hasText === true;
        },
        textRange: {
          get text() {
            const c = entry().committed;
            if (!c || !c.hasText) {
              throw new Error("textRange.text must not be read when hasText is false");
            }
            return c.text;
          },
          set text(v: string) {
            const p = ensurePending(entry());
            p.text = v;
            p.hasText = v.length > 0;
          },
          load() {},
        },
        load() {},
      },
      load() {},
      delete() {
        entry().pending = "delete";
      },
    };
  }

  function makeShapesCollection(sheetName: string) {
    const sheet = getSheet(sheetName);
    let loadedItems: ReturnType<typeof makeProxy>[] = [];
    let pendingLoad = false;

    return {
      get items() {
        return loadedItems;
      },
      load(_props?: string) {
        pendingLoad = true;
      },
      addGeometricShape(geometricShapeType: string) {
        seq += 1;
        const key = `__new_${seq}`;
        const state: ShapeState = {
          name: `Shape${seq}`,
          type: "GeometricShape",
          geometricShapeType,
          left: 0,
          top: 0,
          width: 100,
          height: 100,
          visible: true,
          hasText: false,
          text: "",
        };
        sheet.byKey.set(key, { committed: null, pending: state });
        sheet.order.push(key);
        return makeProxy(sheetName, key);
      },
      addTextBox(text?: string) {
        seq += 1;
        const key = `__new_${seq}`;
        // Omitted text → no content (hasText false). Empty string also hasText false.
        const hasText = typeof text === "string" && text.length > 0;
        const state: ShapeState = {
          name: `TextBox${seq}`,
          type: "GeometricShape",
          geometricShapeType: null,
          left: 0,
          top: 0,
          width: 100,
          height: 40,
          visible: true,
          hasText,
          text: hasText ? text! : "",
        };
        sheet.byKey.set(key, { committed: null, pending: state });
        sheet.order.push(key);
        return makeProxy(sheetName, key);
      },
      getItem(name: string) {
        for (const [key, entry] of sheet.byKey) {
          const live = effective(entry) ?? entry.committed;
          if (live && live.name === name) return makeProxy(sheetName, key);
        }
        // Also resolve by committed name only for post-sync lookups
        for (const [key, entry] of sheet.byKey) {
          if (entry.committed?.name === name) return makeProxy(sheetName, key);
        }
        throw new Error(`missing shape ${name}`);
      },
      _flushLoad() {
        if (!pendingLoad) return;
        pendingLoad = false;
        loadedItems = [];
        for (const key of sheet.order) {
          const entry = sheet.byKey.get(key);
          if (!entry?.committed) continue;
          loadedItems.push(makeProxy(sheetName, key));
        }
      },
      _commit() {
        const nextOrder: string[] = [];
        for (const key of sheet.order) {
          const entry = sheet.byKey.get(key);
          if (!entry) continue;
          if (entry.pending === "delete") {
            sheet.byKey.delete(key);
            continue;
          }
          if (entry.pending !== undefined) {
            entry.committed = clone(entry.pending);
            entry.pending = undefined;
          }
          if (entry.committed) nextOrder.push(key);
        }
        sheet.order = nextOrder;
      },
    };
  }

  const collections = new Map<string, ReturnType<typeof makeShapesCollection>>();
  for (const name of sheets.keys()) {
    collections.set(name, makeShapesCollection(name));
  }

  function makeSheetApi(name: string) {
    return {
      name,
      load() {},
      get shapes() {
        return collections.get(name)!;
      },
      getRange() {
        return { load() {}, address: `${name}!A1` };
      },
    };
  }

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(batch: (context: unknown) => Promise<T>) => {
      const context = {
        workbook: {
          worksheets: {
            getItem(name: string) {
              if (!sheets.has(name)) throw new Error(`missing ${name}`);
              return makeSheetApi(name);
            },
            get items() {
              return [...sheets.keys()].map((n) => makeSheetApi(n));
            },
            load() {},
          },
          load() {},
        },
        async sync() {
          for (const col of collections.values()) {
            col._commit();
          }
          for (const col of collections.values()) {
            col._flushLoad();
          }
        },
      };
      return batch(context);
    },
  };

  return {
    seedNoTextShape(sheetName: string, name: string) {
      const sheet = getSheet(sheetName);
      const state: ShapeState = {
        name,
        type: "GeometricShape",
        geometricShapeType: "Rectangle",
        left: 10,
        top: 20,
        width: 50,
        height: 60,
        visible: true,
        hasText: false,
        text: "",
      };
      sheet.byKey.set(name, { committed: state, pending: undefined });
      if (!sheet.order.includes(name)) sheet.order.push(name);
    },
  };
}
