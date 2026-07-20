/**
 * DV proxy: rule + ignoreBlanks are both sync-gated.
 * list.source may remain Range-like after commit.
 */
import {
  applyDvTamper,
  guessDvType,
  materializeRule,
  type CtxPending,
  type DvState,
  type ValidationFakeOptions,
} from "./officeJsValidationFakeState";

export function makeDvProxy(
  key: string,
  dvs: Map<string, DvState>,
  pending: CtxPending,
  options: {
    keepListSourceAsRangeObject: boolean;
    getTamper: () => ValidationFakeOptions["tamperDvReadback"];
    clearLeavesHostType?: string;
  },
) {
  let localIgnoreBlanks = true;
  let localType: string | null = null;
  let localRule: Record<string, unknown> = {};
  let loaded = false;

  function queueWrite(partial: Partial<DvState> & { rule?: Record<string, unknown> }) {
    const existing = pending.dvWrites.get(key);
    if (existing === "clear") {
      pending.dvWrites.delete(key);
    }
    const cur = pending.dvWrites.get(key);
    if (cur && cur !== "clear") {
      if (partial.rule) cur.rule = partial.rule;
      if (partial.type !== undefined) cur.type = partial.type;
      if (partial.ignoreBlanks !== undefined) cur.ignoreBlanks = partial.ignoreBlanks;
      return;
    }
    pending.dvWrites.set(key, {
      type: partial.type ?? localType,
      ignoreBlanks: partial.ignoreBlanks ?? localIgnoreBlanks,
      rule: partial.rule ?? localRule,
    });
  }

  const dvProxy: {
    type: string | null;
    ignoreBlanks: boolean;
    rule: Record<string, unknown>;
    load: (props?: string) => void;
    clear: () => void;
  } = {
    get type() {
      return loaded ? localType : localType;
    },
    set type(v: string | null) {
      localType = v;
    },
    get ignoreBlanks() {
      return localIgnoreBlanks;
    },
    set ignoreBlanks(v: boolean) {
      localIgnoreBlanks = v;
      queueWrite({ ignoreBlanks: v });
    },
    get rule() {
      return localRule;
    },
    set rule(next: Record<string, unknown>) {
      const copy = materializeRule(next, options.keepListSourceAsRangeObject);
      const type = guessDvType(copy);
      localRule = copy;
      localType = type;
      queueWrite({ type, ignoreBlanks: localIgnoreBlanks, rule: copy });
    },
    load() {
      pending.loads.push(() => {
        loaded = true;
        const committed = dvs.get(key);
        if (!committed) {
          localType = options.clearLeavesHostType ?? null;
          localRule = {};
          localIgnoreBlanks = true;
          // clearLeavesHostType simulates bad clear readback
          if (options.clearLeavesHostType) {
            localType = options.clearLeavesHostType;
          } else {
            localType = null;
          }
          return;
        }
        const view = applyDvTamper(committed, options.getTamper());
        localType = view.type;
        localIgnoreBlanks = view.ignoreBlanks;
        // Preserve Range object references when present.
        localRule = view.rule;
      });
    },
    clear() {
      pending.dvWrites.set(key, "clear");
      localType = null;
      localRule = {};
    },
  };

  return dvProxy;
}
