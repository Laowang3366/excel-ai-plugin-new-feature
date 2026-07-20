/** CF proxy builders: cellValue whole-object rule assign vs custom ClientObject formula. */
import type { CfState, CtxPending } from "./officeJsValidationFakeState";

export function makeCfProxy(key: string, state: CfState, pending: CtxPending) {
  return {
    get id() {
      return state.id;
    },
    get type() {
      return state.type;
    },
    get cellValue() {
      if (state.type !== "CellValue") return undefined;
      return {
        get rule() {
          return state.cellRule ?? { formula1: "", operator: "EqualTo" };
        },
        set rule(next: { formula1: string; formula2?: string; operator: string }) {
          // Whole-object assign (plain data) — mirror Office.js ConditionalCellValueRule.
          state.cellRule = { ...next };
          pending.cfPatches.push({ key, id: state.id, patch: { cellRule: { ...next } } });
        },
        format: {
          fill: {
            get color() {
              return state.cellFill ?? "";
            },
            set color(v: string) {
              state.cellFill = v;
              pending.cfPatches.push({ key, id: state.id, patch: { cellFill: v } });
            },
          },
          font: {
            get color() {
              return state.cellFont ?? "";
            },
            set color(v: string) {
              state.cellFont = v;
              pending.cfPatches.push({ key, id: state.id, patch: { cellFont: v } });
            },
          },
        },
      };
    },
    get custom() {
      if (state.type !== "Custom") return undefined;
      return {
        // ClientObject: only formula property is queueable (not whole rule replace).
        rule: {
          get formula() {
            return state.customFormula ?? "";
          },
          set formula(v: string) {
            state.customFormula = v;
            pending.cfPatches.push({
              key,
              id: state.id,
              patch: { customFormula: v },
            });
          },
        },
        format: {
          fill: {
            set color(v: string) {
              state.customFill = v;
              pending.cfPatches.push({ key, id: state.id, patch: { customFill: v } });
            },
          },
          font: {
            set color(v: string) {
              state.customFont = v;
              pending.cfPatches.push({ key, id: state.id, patch: { customFont: v } });
            },
          },
        },
      };
    },
    load() {},
    delete() {
      pending.cfDeletes.push({ key, id: state.id });
    },
  };
}

export function makeConditionalFormatsApi(opts: {
  key: string;
  cfs: Map<string, CfState[]>;
  pending: CtxPending;
  nextSeq: () => number;
  getItemProxies: () => ReturnType<typeof makeCfProxy>[];
  setItemProxies: (items: ReturnType<typeof makeCfProxy>[]) => void;
}) {
  const { key, cfs, pending, nextSeq, getItemProxies, setItemProxies } = opts;
  return {
    get items() {
      return getItemProxies();
    },
    load() {
      pending.loads.push(() => {
        const list = cfs.get(key) ?? [];
        setItemProxies(list.map((s) => makeCfProxy(key, { ...s }, pending)));
      });
    },
    add(type: string) {
      const state: CfState = {
        id: `cf_${nextSeq()}`,
        type,
        cellRule:
          type === "CellValue" ? { formula1: "", operator: "EqualTo" } : undefined,
        customFormula: type === "Custom" ? "" : undefined,
      };
      pending.cfAdds.push({ key, state });
      return makeCfProxy(key, state, pending);
    },
    getItem(id: string) {
      const list = cfs.get(key) ?? [];
      const hit = list.find((s) => s.id === id);
      if (!hit) throw new Error(`ItemNotFound: ${id}`);
      return makeCfProxy(key, { ...hit }, pending);
    },
  };
}
