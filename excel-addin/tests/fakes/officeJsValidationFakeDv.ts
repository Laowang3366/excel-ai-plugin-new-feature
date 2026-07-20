/** DV proxy: rule assign queues until sync; load materializes committed state. */
import {
  guessDvType,
  materializeRule,
  type CtxPending,
  type DvState,
} from "./officeJsValidationFakeState";

export function makeDvProxy(key: string, dvs: Map<string, DvState>, pending: CtxPending) {
  const dvProxy: {
    type: string | null;
    ignoreBlanks: boolean;
    _rule: Record<string, unknown>;
    load: (props?: string) => void;
    clear: () => void;
  } = {
    type: null,
    ignoreBlanks: true,
    _rule: {},
    load() {
      pending.loads.push(() => {
        const committed = dvs.get(key);
        if (!committed) {
          dvProxy.type = null;
          dvProxy._rule = {};
          dvProxy.ignoreBlanks = true;
          return;
        }
        dvProxy.type = committed.type;
        dvProxy._rule = JSON.parse(JSON.stringify(committed.rule));
        dvProxy.ignoreBlanks = committed.ignoreBlanks;
      });
    },
    clear() {
      pending.dvWrites.set(key, "clear");
    },
  };

  Object.defineProperty(dvProxy, "rule", {
    get() {
      return dvProxy._rule;
    },
    set(next: Record<string, unknown>) {
      const copy = materializeRule(next);
      const type = guessDvType(copy);
      pending.dvWrites.set(key, {
        type,
        ignoreBlanks: dvProxy.ignoreBlanks,
        rule: copy,
      });
      dvProxy._rule = copy;
      dvProxy.type = type;
    },
    enumerable: true,
    configurable: true,
  });

  return dvProxy;
}
