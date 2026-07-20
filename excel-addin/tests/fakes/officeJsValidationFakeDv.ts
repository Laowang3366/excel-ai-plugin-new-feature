/**
 * DV proxy: rule + ignoreBlanks + errorAlert/prompt are sync-gated.
 * list.source may remain Range-like after commit.
 */
import {
  applyDvTamper,
  guessDvType,
  materializeRule,
  rehydrateDvRule,
  type CtxPending,
  type DvAlertState,
  type DvPromptState,
  type DvState,
  type ValidationFakeOptions,
} from "./officeJsValidationFakeState";

function makeAlertBag(initial?: DvAlertState) {
  return {
    showAlert: initial?.showAlert,
    style: initial?.style,
    title: initial?.title,
    message: initial?.message,
  };
}

function makePromptBag(initial?: DvPromptState) {
  return {
    showPrompt: initial?.showPrompt,
    title: initial?.title,
    message: initial?.message,
  };
}

export function makeDvProxy(
  key: string,
  dvs: Map<string, DvState>,
  pending: CtxPending,
  options: {
    keepListSourceAsRangeObject: boolean;
    getTamper: () => ValidationFakeOptions["tamperDvReadback"];
    clearLeavesHostType?: string;
    missingDvErrorAlert?: boolean;
    missingDvPrompt?: boolean;
    missingDvErrorAlertFields?: boolean;
  },
) {
  let localIgnoreBlanks = true;
  let localType: string | null = null;
  let localRule: Record<string, unknown> = {};
  let localErrorAlert = makeAlertBag();
  let localPrompt = makePromptBag();

  function queueWrite(
    partial: Partial<DvState> & {
      rule?: Record<string, unknown>;
      touchErrorAlert?: boolean;
      touchPrompt?: boolean;
    },
  ) {
    const existing = pending.dvWrites.get(key);
    if (existing === "clear") {
      pending.dvWrites.delete(key);
    }
    const cur = pending.dvWrites.get(key);
    if (cur && cur !== "clear") {
      if (partial.rule) cur.rule = partial.rule;
      if (partial.type !== undefined) cur.type = partial.type;
      if (partial.ignoreBlanks !== undefined) cur.ignoreBlanks = partial.ignoreBlanks;
      if (partial.touchErrorAlert) cur.errorAlert = { ...localErrorAlert };
      if (partial.touchPrompt) cur.prompt = { ...localPrompt };
      return;
    }
    const next: DvState = {
      type: partial.type ?? localType,
      ignoreBlanks: partial.ignoreBlanks ?? localIgnoreBlanks,
      rule: partial.rule ?? localRule,
    };
    // Only attach metadata keys when explicitly touched — omit keeps host values on sync.
    if (partial.touchErrorAlert) next.errorAlert = { ...localErrorAlert };
    if (partial.touchPrompt) next.prompt = { ...localPrompt };
    pending.dvWrites.set(key, next);
  }

  function touchAlert() {
    queueWrite({ touchErrorAlert: true });
  }
  function touchPrompt() {
    queueWrite({ touchPrompt: true });
  }

  const errorAlertProxy = options.missingDvErrorAlert
    ? undefined
    : options.missingDvErrorAlertFields
      ? ({} as DvAlertState)
      : {
          get showAlert() {
            return localErrorAlert.showAlert;
          },
          set showAlert(v: boolean | undefined) {
            localErrorAlert.showAlert = v;
            touchAlert();
          },
          get style() {
            return localErrorAlert.style;
          },
          set style(v: string | undefined) {
            localErrorAlert.style = v;
            touchAlert();
          },
          get title() {
            return localErrorAlert.title;
          },
          set title(v: string | undefined) {
            localErrorAlert.title = v;
            touchAlert();
          },
          get message() {
            return localErrorAlert.message;
          },
          set message(v: string | undefined) {
            localErrorAlert.message = v;
            touchAlert();
          },
        };

  const promptProxy = options.missingDvPrompt
    ? undefined
    : {
        get showPrompt() {
          return localPrompt.showPrompt;
        },
        set showPrompt(v: boolean | undefined) {
          localPrompt.showPrompt = v;
          touchPrompt();
        },
        get title() {
          return localPrompt.title;
        },
        set title(v: string | undefined) {
          localPrompt.title = v;
          touchPrompt();
        },
        get message() {
          return localPrompt.message;
        },
        set message(v: string | undefined) {
          localPrompt.message = v;
          touchPrompt();
        },
      };

  const dvProxy: {
    type: string | null;
    ignoreBlanks: boolean;
    rule: Record<string, unknown>;
    errorAlert?: typeof errorAlertProxy;
    prompt?: typeof promptProxy;
    load: (props?: string) => void;
    clear: () => void;
  } = {
    get type() {
      return localType;
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
        const committed = dvs.get(key);
        if (!committed) {
          localRule = {};
          localIgnoreBlanks = true;
          localErrorAlert = makeAlertBag();
          localPrompt = makePromptBag();
          if (options.clearLeavesHostType) {
            localType = options.clearLeavesHostType;
          } else {
            localType = "None";
          }
          return;
        }
        const view = applyDvTamper(committed, options.getTamper());
        localType = view.type;
        localIgnoreBlanks = view.ignoreBlanks;
        localRule = rehydrateDvRule(
          view.rule,
          pending,
          options.keepListSourceAsRangeObject,
        );
        localErrorAlert = makeAlertBag(view.errorAlert);
        localPrompt = makePromptBag(view.prompt);
      });
    },
    clear() {
      pending.dvWrites.set(key, "clear");
      localType = null;
      localRule = {};
      localErrorAlert = makeAlertBag();
      localPrompt = makePromptBag();
    },
  };

  if (errorAlertProxy !== undefined) {
    Object.defineProperty(dvProxy, "errorAlert", {
      enumerable: true,
      configurable: true,
      get: () => errorAlertProxy,
    });
  }
  if (promptProxy !== undefined) {
    Object.defineProperty(dvProxy, "prompt", {
      enumerable: true,
      configurable: true,
      get: () => promptProxy,
    });
  }

  return dvProxy;
}
