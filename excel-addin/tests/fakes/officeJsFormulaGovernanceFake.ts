/**
 * Sync-gated multi-sheet fake for formula governance.
 */
type SheetState = {
  name: string;
  visibility: string;
  formulas: string[][];
  values: unknown[][];
  numberFormat: string[][];
};

export function installFormulaGovernanceExcel(options?: {
  sheets?: Array<{ name: string; formulas: string[][]; values?: unknown[][] }>;
  excelApi12?: boolean;
}) {
  const excelApi12 = options?.excelApi12 !== false;
  const sheetList: SheetState[] = (options?.sheets ?? [
    {
      name: "Sheet1",
      formulas: [
        ["=B1", "=1+#REF!"],
        ["=10", "x"],
      ],
      values: [
        [1, "#REF!"],
        [10, "x"],
      ],
    },
  ]).map((s) => ({
    name: s.name,
    visibility: "Visible",
    formulas: s.formulas.map((row) => [...row]),
    values: (s.values ?? s.formulas.map((row) => row.map(() => null))).map((row) => [...row]),
    numberFormat: s.formulas.map((row) => row.map(() => "General")),
  }));

  function findSheet(name: string): SheetState {
    const hit = sheetList.find((s) => s.name === name);
    if (!hit) throw new Error(`Worksheet ${name} not found`);
    return hit;
  }

  function parseA1(address: string): { r0: number; c0: number; r1: number; c1: number } {
    const bare = address.includes("!") ? address.slice(address.lastIndexOf("!") + 1) : address;
    const clean = bare.replace(/\$/g, "").toUpperCase();
    const parts = clean.split(":");
    const one = (a: string) => {
      const m = /^([A-Z]+)(\d+)$/.exec(a);
      if (!m) throw new Error(`bad address ${a}`);
      let col = 0;
      for (const ch of m[1]!) col = col * 26 + (ch.charCodeAt(0) - 64);
      return { r: Number(m[2]) - 1, c: col - 1 };
    };
    const a = one(parts[0]!);
    const b = parts[1] ? one(parts[1]!) : a;
    return { r0: a.r, c0: a.c, r1: b.r, c1: b.c };
  }

  function toA1(r: number, c: number): string {
    let n = c + 1;
    let label = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      label = String.fromCharCode(65 + rem) + label;
      n = Math.floor((n - 1) / 26);
    }
    return `${label}${r + 1}`;
  }

  function ensureSize(sheet: SheetState, r: number, c: number) {
    while (sheet.formulas.length <= r) {
      sheet.formulas.push([]);
      sheet.values.push([]);
      sheet.numberFormat.push([]);
    }
    while ((sheet.formulas[r]?.length ?? 0) <= c) {
      sheet.formulas[r]!.push("");
      sheet.values[r]!.push(null);
      sheet.numberFormat[r]!.push("General");
    }
  }

  const allRanges: Array<{ commit: () => void }> = [];

  function makeRange(sheet: SheetState, address: string) {
    const box = parseA1(address);
    const loaded: Record<string, unknown> = {};
    let pendingFormulas: string[][] | undefined;
    let pendingValues: unknown[][] | undefined;
    let pendingNf: string[][] | undefined;

    const commit = () => {
      if (pendingFormulas) {
        for (let r = 0; r < pendingFormulas.length; r++) {
          for (let c = 0; c < (pendingFormulas[r]?.length ?? 0); c++) {
            ensureSize(sheet, box.r0 + r, box.c0 + c);
            sheet.formulas[box.r0 + r]![box.c0 + c] = pendingFormulas[r]![c] ?? "";
          }
        }
        pendingFormulas = undefined;
      }
      if (pendingValues) {
        for (let r = 0; r < pendingValues.length; r++) {
          for (let c = 0; c < (pendingValues[r]?.length ?? 0); c++) {
            ensureSize(sheet, box.r0 + r, box.c0 + c);
            sheet.values[box.r0 + r]![box.c0 + c] = pendingValues[r]![c];
            sheet.formulas[box.r0 + r]![box.c0 + c] = "";
          }
        }
        pendingValues = undefined;
      }
      if (pendingNf) {
        for (let r = 0; r < pendingNf.length; r++) {
          for (let c = 0; c < (pendingNf[r]?.length ?? 0); c++) {
            ensureSize(sheet, box.r0 + r, box.c0 + c);
            sheet.numberFormat[box.r0 + r]![box.c0 + c] = pendingNf[r]![c] ?? "General";
          }
        }
        pendingNf = undefined;
      }
    };
    allRanges.push({ commit });

    return {
      isNullObject: false,
      get address() {
        if (!("address" in loaded)) throw new Error("address not loaded");
        return loaded.address as string;
      },
      get formulas() {
        if (!("formulas" in loaded)) throw new Error("formulas not loaded");
        return loaded.formulas as string[][];
      },
      set formulas(v: string[][]) {
        pendingFormulas = v;
      },
      get values() {
        if (!("values" in loaded)) throw new Error("values not loaded");
        return loaded.values as unknown[][];
      },
      set values(v: unknown[][]) {
        pendingValues = v;
      },
      get numberFormat() {
        if (!("numberFormat" in loaded)) throw new Error("numberFormat not loaded");
        return loaded.numberFormat as string[][];
      },
      set numberFormat(v: string[][] | string) {
        pendingNf = Array.isArray(v) ? v : [[String(v)]];
      },
      get rowCount() {
        return box.r1 - box.r0 + 1;
      },
      get columnCount() {
        return box.c1 - box.c0 + 1;
      },
      format: {
        protection: {
          locked: false,
          load() {},
        },
      },
      load(props: string) {
        const set = new Set(props.split(",").map((p) => p.trim()));
        if (set.has("address") || set.has("isNullObject")) {
          loaded.address = `${sheet.name}!${toA1(box.r0, box.c0)}:${toA1(box.r1, box.c1)}`;
          loaded.isNullObject = false;
        }
        if (set.has("formulas")) {
          const m: string[][] = [];
          for (let r = box.r0; r <= box.r1; r++) {
            const row: string[] = [];
            for (let c = box.c0; c <= box.c1; c++) {
              ensureSize(sheet, r, c);
              row.push(sheet.formulas[r]![c] ?? "");
            }
            m.push(row);
          }
          loaded.formulas = m;
        }
        if (set.has("values")) {
          const m: unknown[][] = [];
          for (let r = box.r0; r <= box.r1; r++) {
            const row: unknown[] = [];
            for (let c = box.c0; c <= box.c1; c++) {
              ensureSize(sheet, r, c);
              row.push(sheet.values[r]![c] ?? null);
            }
            m.push(row);
          }
          loaded.values = m;
        }
        if (set.has("numberFormat")) {
          const m: string[][] = [];
          for (let r = box.r0; r <= box.r1; r++) {
            const row: string[] = [];
            for (let c = box.c0; c <= box.c1; c++) {
              ensureSize(sheet, r, c);
              row.push(sheet.numberFormat[r]![c] ?? "General");
            }
            m.push(row);
          }
          loaded.numberFormat = m;
        }
        if (set.has("rowCount")) loaded.rowCount = box.r1 - box.r0 + 1;
        if (set.has("columnCount")) loaded.columnCount = box.c1 - box.c0 + 1;
      },
      getCell(row: number, col: number) {
        return makeRange(sheet, toA1(box.r0 + row, box.c0 + col));
      },
      clear() {
        for (let r = box.r0; r <= box.r1; r++) {
          for (let c = box.c0; c <= box.c1; c++) {
            ensureSize(sheet, r, c);
            sheet.formulas[r]![c] = "";
            sheet.values[r]![c] = null;
          }
        }
      },
    };
  }

  function makeWorksheet(state: SheetState) {
    return {
      get name() {
        return state.name;
      },
      set name(v: string) {
        state.name = v;
      },
      get visibility() {
        return state.visibility;
      },
      set visibility(v: string) {
        state.visibility = v;
      },
      load(_props: string) {},
      getRange(address: string) {
        return makeRange(state, address);
      },
      getUsedRangeOrNullObject() {
        let maxR = 0;
        let maxC = 0;
        let any = false;
        for (let r = 0; r < state.formulas.length; r++) {
          for (let c = 0; c < (state.formulas[r]?.length ?? 0); c++) {
            if (state.formulas[r]![c] || state.values[r]![c] != null) {
              maxR = Math.max(maxR, r);
              maxC = Math.max(maxC, c);
              any = true;
            }
          }
        }
        if (!any) {
          const empty = makeRange(state, "A1");
          (empty as { isNullObject: boolean }).isNullObject = true;
          return empty;
        }
        return makeRange(state, `A1:${toA1(maxR, maxC)}`);
      },
      protection: { protected: false, load() {}, unprotect() {}, protect() {} },
    };
  }

  const worksheets = {
    items: [] as ReturnType<typeof makeWorksheet>[],
    load() {
      this.items = sheetList.map((s) => makeWorksheet(s));
    },
    getItem(name: string) {
      findSheet(name);
      return makeWorksheet(findSheet(name));
    },
    add(name: string) {
      const state: SheetState = {
        name,
        visibility: "Visible",
        formulas: [[""]],
        values: [[null]],
        numberFormat: [["General"]],
      };
      sheetList.push(state);
      const ws = makeWorksheet(state);
      this.items.push(ws);
      return ws;
    },
  };

  async function run(fn: (ctx: unknown) => Promise<unknown>) {
    const context = {
      workbook: { worksheets, load() {} },
      async sync() {
        for (const r of allRanges) r.commit();
      },
    };
    return fn(context);
  }

  (globalThis as unknown as { Excel: { run: typeof run } }).Excel = { run };
  (globalThis as unknown as {
    Office: { context: { requirements: { isSetSupported: (s: string, v: string) => boolean } } };
  }).Office = {
    context: {
      requirements: {
        isSetSupported(set: string, version: string) {
          if (set === "ExcelApi" && version === "1.2") return excelApi12;
          return true;
        },
      },
    },
  };
  (globalThis as unknown as { window: typeof globalThis }).window = globalThis;

  return {
    sheets: () => sheetList,
    formulas: (name = "Sheet1") => findSheet(name).formulas,
    values: (name = "Sheet1") => findSheet(name).values,
    backupSheet: () => sheetList.find((s) => s.name.startsWith("_WenggeFormulaBackup")) ?? null,
  };
}
