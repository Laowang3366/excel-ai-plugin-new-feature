/**
 * Sync-gated multi-sheet fake for formula governance.
 * - values assignment must match Range row/col dimensions (throws otherwise)
 * - formula-like strings without text format / apostrophe are "evaluated" (lose leading =)
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
  /** When false, worksheet.visibility assignment is ignored / not VeryHidden. */
  supportVeryHidden?: boolean;
}) {
  const excelApi12 = options?.excelApi12 !== false;
  const supportVeryHidden = options?.supportVeryHidden !== false;
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

  function looksFormulaLike(v: unknown): v is string {
    return typeof v === "string" && (v.startsWith("=") || v.startsWith("+") || v.startsWith("-") || v.startsWith("@"));
  }

  function makeRange(sheet: SheetState, address: string) {
    const box = parseA1(address);
    const rowCount = box.r1 - box.r0 + 1;
    const colCount = box.c1 - box.c0 + 1;
    const loaded: Record<string, unknown> = {};
    let pendingFormulas: string[][] | undefined;
    let pendingValues: unknown[][] | undefined;
    let pendingNf: string[][] | undefined;

    const commit = () => {
      if (pendingNf) {
        for (let r = 0; r < pendingNf.length; r++) {
          for (let c = 0; c < (pendingNf[r]?.length ?? 0); c++) {
            ensureSize(sheet, box.r0 + r, box.c0 + c);
            sheet.numberFormat[box.r0 + r]![box.c0 + c] = pendingNf[r]![c] ?? "General";
          }
        }
        pendingNf = undefined;
      }
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
            let v = pendingValues[r]![c];
            const nf = sheet.numberFormat[box.r0 + r]![box.c0 + c] ?? "General";
            const isText = nf === "@" || String(nf).includes("@");
            if (typeof v === "string" && v.startsWith("'")) {
              // apostrophe text marker — store without evaluating; strip one apostrophe on read path stored value
              sheet.values[box.r0 + r]![box.c0 + c] = v; // keep apostrophe in store for decode
              sheet.formulas[box.r0 + r]![box.c0 + c] = "";
            } else if (looksFormulaLike(v) && !isText) {
              // Simulated evaluation: lose formula text (backup corruption risk)
              sheet.values[box.r0 + r]![box.c0 + c] = "EVALUATED";
              sheet.formulas[box.r0 + r]![box.c0 + c] = String(v);
            } else {
              sheet.values[box.r0 + r]![box.c0 + c] = v;
              sheet.formulas[box.r0 + r]![box.c0 + c] = "";
            }
          }
        }
        pendingValues = undefined;
      }
    };
    allRanges.push({ commit });

    function assertMatrixShape(matrix: unknown[][], label: string) {
      if (!Array.isArray(matrix) || matrix.length !== rowCount) {
        throw new Error(
          `Dimension mismatch on ${label}: range ${rowCount}x${colCount}, got rows=${Array.isArray(matrix) ? matrix.length : "?"}`,
        );
      }
      for (let r = 0; r < matrix.length; r++) {
        if (!Array.isArray(matrix[r]) || matrix[r]!.length !== colCount) {
          throw new Error(
            `Dimension mismatch on ${label}: range ${rowCount}x${colCount}, row ${r} cols=${matrix[r]?.length ?? "?"}`,
          );
        }
      }
    }

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
        assertMatrixShape(v, "formulas");
        pendingFormulas = v;
      },
      get values() {
        if (!("values" in loaded)) throw new Error("values not loaded");
        return loaded.values as unknown[][];
      },
      set values(v: unknown[][]) {
        assertMatrixShape(v, "values");
        pendingValues = v;
      },
      get numberFormat() {
        if (!("numberFormat" in loaded)) throw new Error("numberFormat not loaded");
        return loaded.numberFormat as string[][];
      },
      set numberFormat(v: string[][] | string) {
        const matrix = Array.isArray(v) ? v : Array.from({ length: rowCount }, () => Array(colCount).fill(String(v)));
        assertMatrixShape(matrix as unknown[][], "numberFormat");
        pendingNf = matrix as string[][];
      },
      get rowCount() {
        return rowCount;
      },
      get columnCount() {
        return colCount;
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
        if (set.has("formulas") || set.has("formulasR1C1")) {
          const m: string[][] = [];
          for (let r = box.r0; r <= box.r1; r++) {
            const row: string[] = [];
            for (let c = box.c0; c <= box.c1; c++) {
              ensureSize(sheet, r, c);
              row.push(sheet.formulas[r]![c] ?? "");
            }
            m.push(row);
          }
          if (set.has("formulas")) loaded.formulas = m;
          if (set.has("formulasR1C1")) loaded.formulasR1C1 = m.map((row) => row.map((f) => f));
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
        if (set.has("rowCount")) loaded.rowCount = rowCount;
        if (set.has("columnCount")) loaded.columnCount = colCount;
      },
      get formulasR1C1() {
        return (loaded.formulasR1C1 as string[][]) ?? [];
      },
      getCell(row: number, col: number) {
        return makeRange(sheet, toA1(box.r0 + row, box.c0 + col));
      },
      getSpillingToRange() {
        // no spill in default fake
        const empty = makeRange(sheet, toA1(box.r0, box.c0));
        (empty as { isNullObject: boolean }).isNullObject = true;
        return empty;
      },
      clear() {
        for (let r = box.r0; r <= box.r1; r++) {
          for (let c = box.c0; c <= box.c1; c++) {
            ensureSize(sheet, r, c);
            sheet.formulas[r]![c] = "";
            sheet.values[r]![c] = null;
            sheet.numberFormat[r]![c] = "General";
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
        if (!supportVeryHidden && String(v).toLowerCase().includes("veryhidden")) {
          state.visibility = "Visible";
          return;
        }
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
        for (let r = 0; r < Math.max(state.formulas.length, state.values.length); r++) {
          const frow = state.formulas[r] ?? [];
          const vrow = state.values[r] ?? [];
          const len = Math.max(frow.length, vrow.length);
          for (let c = 0; c < len; c++) {
            if (frow[c] || vrow[c] != null) {
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
