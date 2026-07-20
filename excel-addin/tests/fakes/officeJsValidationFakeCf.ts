/** CF proxy: PropertyNotLoaded until load+sync; cellValue rule whole-object vs custom ClientObject. */
import {
  applyCfTamper,
  type CfState,
  type CtxPending,
  type ValidationFakeOptions,
} from "./officeJsValidationFakeState";

export function makeCfProxy(
  key: string,
  state: CfState,
  pending: CtxPending,
  options: {
    /** When true, id/type require load (add-returned proxy). Collection items start loaded. */
    requireLoadForIdType: boolean;
    getTamper: () => ValidationFakeOptions["tamperCfReadback"];
  },
) {
  const loaded = new Set<string>();
  const pendingProps = new Set<string>();
  if (!options.requireLoadForIdType) {
    loaded.add("id");
    loaded.add("type");
    loaded.add("rule");
    loaded.add("formula");
    loaded.add("fillColor");
    loaded.add("fontColor");
  }

  function ensure(prop: string): void {
    if (!loaded.has(prop)) {
      throw new Error(`PropertyNotLoaded:${prop}`);
    }
  }

  function viewState(): CfState {
    return applyCfTamper(state, options.getTamper());
  }

  return {
    get id() {
      ensure("id");
      return state.id;
    },
    get type() {
      ensure("type");
      return state.type;
    },
    get cellValue() {
      if (state.type !== "CellValue") return undefined;
      return {
        get rule() {
          ensure("rule");
          const v = viewState();
          return v.cellRule ?? { formula1: "", operator: "EqualTo" };
        },
        set rule(next: { formula1: string; formula2?: string; operator: string }) {
          state.cellRule = { ...next };
          pending.cfPatches.push({ key, id: state.id, patch: { cellRule: { ...next } } });
        },
        format: {
          fill: {
            get color() {
              ensure("fillColor");
              return viewState().cellFill ?? "";
            },
            set color(v: string) {
              state.cellFill = v;
              pending.cfPatches.push({ key, id: state.id, patch: { cellFill: v } });
            },
            load(props: string) {
              if (props.includes("color")) pendingProps.add("fillColor");
              pending.loads.push(() => {
                for (const p of pendingProps) loaded.add(p);
              });
            },
          },
          font: {
            get color() {
              ensure("fontColor");
              return viewState().cellFont ?? "";
            },
            set color(v: string) {
              state.cellFont = v;
              pending.cfPatches.push({ key, id: state.id, patch: { cellFont: v } });
            },
            load(props: string) {
              if (props.includes("color")) pendingProps.add("fontColor");
              pending.loads.push(() => {
                for (const p of pendingProps) loaded.add(p);
              });
            },
          },
        },
        load(props: string) {
          if (props.includes("rule")) pendingProps.add("rule");
          pending.loads.push(() => {
            for (const p of pendingProps) loaded.add(p);
          });
        },
      };
    },
    get custom() {
      if (state.type !== "Custom") return undefined;
      return {
        rule: {
          get formula() {
            ensure("formula");
            return viewState().customFormula ?? "";
          },
          set formula(v: string) {
            state.customFormula = v;
            pending.cfPatches.push({
              key,
              id: state.id,
              patch: { customFormula: v },
            });
          },
          load(props: string) {
            if (props.includes("formula")) pendingProps.add("formula");
            pending.loads.push(() => {
              for (const p of pendingProps) loaded.add(p);
            });
          },
        },
        format: {
          fill: {
            get color() {
              ensure("fillColor");
              return viewState().customFill ?? "";
            },
            set color(v: string) {
              state.customFill = v;
              pending.cfPatches.push({ key, id: state.id, patch: { customFill: v } });
            },
            load(props: string) {
              if (props.includes("color")) pendingProps.add("fillColor");
              pending.loads.push(() => {
                for (const p of pendingProps) loaded.add(p);
              });
            },
          },
          font: {
            get color() {
              ensure("fontColor");
              return viewState().customFont ?? "";
            },
            set color(v: string) {
              state.customFont = v;
              pending.cfPatches.push({ key, id: state.id, patch: { customFont: v } });
            },
            load(props: string) {
              if (props.includes("color")) pendingProps.add("fontColor");
              pending.loads.push(() => {
                for (const p of pendingProps) loaded.add(p);
              });
            },
          },
        },
      };
    },
    load(props: string) {
      for (const p of props.split(",").map((s) => s.trim()).filter(Boolean)) {
        pendingProps.add(p);
      }
      pending.loads.push(() => {
        for (const p of pendingProps) loaded.add(p);
      });
    },
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
  getTamper: () => ValidationFakeOptions["tamperCfReadback"];
}) {
  const { key, cfs, pending, nextSeq, getItemProxies, setItemProxies, getTamper } = opts;
  return {
    get items() {
      return getItemProxies();
    },
    load() {
      pending.loads.push(() => {
        const list = cfs.get(key) ?? [];
        setItemProxies(
          list.map((s) =>
            makeCfProxy(key, { ...s }, pending, {
              requireLoadForIdType: false,
              getTamper,
            }),
          ),
        );
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
      // Add-returned proxy requires load("id,type") before reading id/type.
      return makeCfProxy(key, state, pending, {
        requireLoadForIdType: true,
        getTamper,
      });
    },
    getItem(id: string) {
      const list = cfs.get(key) ?? [];
      const hit = list.find((s) => s.id === id);
      if (!hit) throw new Error(`ItemNotFound: ${id}`);
      return makeCfProxy(key, { ...hit }, pending, {
        requireLoadForIdType: false,
        getTamper,
      });
    },
  };
}
