/**
 * Minimal Excel.run fake exercising Range.conditionalFormats + Range.dataValidation
 * with the official property shapes (rule nested; custom.rule.formula string).
 */
export function installValidationExcel() {
  type CfState = {
    id: string;
    type: string;
    cellValue?: {
      rule: { formula1: string; formula2?: string; operator: string };
      format: { fill: { color: string }; font: { color: string } };
    };
    custom?: {
      rule: { formula: string };
      format: { fill: { color: string }; font: { color: string } };
    };
  };

  type DvState = {
    type: string | null;
    ignoreBlanks: boolean;
    rule: {
      list?: { inCellDropDown?: boolean; source?: string };
      wholeNumber?: {
        formula1?: string | number;
        formula2?: string | number;
        operator?: string;
      };
    };
  };

  const cfsByRange = new Map<string, CfState[]>();
  const dvByRange = new Map<string, DvState>();
  let cfSeq = 0;

  function rangeKey(sheet: string, address: string) {
    return `${sheet}!${address.toUpperCase()}`;
  }

  function makeCf(state: CfState) {
    return {
      get id() {
        return state.id;
      },
      get type() {
        return state.type;
      },
      get cellValue() {
        return state.cellValue;
      },
      get custom() {
        return state.custom;
      },
      load() {},
      delete() {
        for (const [key, list] of cfsByRange) {
          cfsByRange.set(
            key,
            list.filter((item) => item.id !== state.id),
          );
        }
      },
    };
  }

  function makeDv(key: string) {
    if (!dvByRange.has(key)) {
      dvByRange.set(key, { type: null, ignoreBlanks: true, rule: {} });
    }
    const state = () => dvByRange.get(key)!;
    return {
      get type() {
        return state().type;
      },
      get ignoreBlanks() {
        return state().ignoreBlanks;
      },
      set ignoreBlanks(v: boolean) {
        state().ignoreBlanks = v;
      },
      get rule() {
        return state().rule;
      },
      set rule(next: DvState["rule"]) {
        state().rule = next;
        if (next.list) state().type = "List";
        else if (next.wholeNumber) state().type = "WholeNumber";
        else state().type = null;
      },
      load() {},
      clear() {
        dvByRange.set(key, { type: null, ignoreBlanks: true, rule: {} });
      },
    };
  }

  function makeRange(sheetName: string, address: string) {
    const key = rangeKey(sheetName, address);
    if (!cfsByRange.has(key)) cfsByRange.set(key, []);
    return {
      address: `${sheetName}!${address}`,
      load() {},
      conditionalFormats: {
        get items() {
          return (cfsByRange.get(key) ?? []).map(makeCf);
        },
        load() {},
        add(type: string) {
          const id = `cf-${++cfSeq}`;
          const isCustom = String(type).toLowerCase().includes("custom");
          const state: CfState = {
            id,
            type: isCustom ? "Custom" : "CellValue",
            cellValue: isCustom
              ? undefined
              : {
                  rule: { formula1: "", operator: "GreaterThan" },
                  format: { fill: { color: "" }, font: { color: "" } },
                },
            custom: isCustom
              ? {
                  rule: { formula: "" },
                  format: { fill: { color: "" }, font: { color: "" } },
                }
              : undefined,
          };
          cfsByRange.get(key)!.push(state);
          return makeCf(state);
        },
        getItem(id: string) {
          const found = (cfsByRange.get(key) ?? []).find((item) => item.id === id);
          if (!found) throw new Error(`missing cf ${id}`);
          return makeCf(found);
        },
      },
      dataValidation: makeDv(key),
    };
  }

  const context = {
    workbook: {
      worksheets: {
        getItem(name: string) {
          return {
            getRange(address: string) {
              return makeRange(name, address);
            },
          };
        },
      },
    },
    async sync() {},
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getCfCount(sheet: string, address: string) {
      return cfsByRange.get(rangeKey(sheet, address))?.length ?? 0;
    },
    getCustomFormula(sheet: string, address: string, id?: string): string | undefined {
      const list = cfsByRange.get(rangeKey(sheet, address)) ?? [];
      const item = id ? list.find((cf) => cf.id === id) : list.find((cf) => cf.custom);
      return item?.custom?.rule.formula;
    },
    getDv(sheet: string, address: string) {
      return dvByRange.get(rangeKey(sheet, address));
    },
  };
}
