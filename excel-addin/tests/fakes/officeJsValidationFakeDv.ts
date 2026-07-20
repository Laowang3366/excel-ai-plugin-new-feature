/**
 * DV proxy: ClientObject-like — PropertyNotLoaded until load+sync;
 * errorAlert/prompt only whole-object setters queue writes;
 * nested mutation of loaded plain snapshots does NOT write.
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

export type DvWriteCounts = {
  rule: number;
  ignoreBlanks: number;
  errorAlert: number;
  prompt: number;
};

const DEFAULT_ERROR: Required<DvAlertState> = {
  showAlert: true,
  style: "Stop",
  title: "",
  message: "",
};

const DEFAULT_PROMPT: Required<DvPromptState> = {
  showPrompt: false,
  title: "",
  message: "",
};

function cloneAlert(a: DvAlertState): Required<DvAlertState> {
  return {
    showAlert: a.showAlert ?? DEFAULT_ERROR.showAlert,
    style: a.style ?? DEFAULT_ERROR.style,
    title: a.title ?? DEFAULT_ERROR.title,
    message: a.message ?? DEFAULT_ERROR.message,
  };
}

function clonePrompt(p: DvPromptState): Required<DvPromptState> {
  return {
    showPrompt: p.showPrompt ?? DEFAULT_PROMPT.showPrompt,
    title: p.title ?? DEFAULT_PROMPT.title,
    message: p.message ?? DEFAULT_PROMPT.message,
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
    poisonSurface?: ValidationFakeOptions["poisonSurface"];
    writeCounts: DvWriteCounts;
    recordLoadProps: (props: string) => void;
  },
) {
  let localIgnoreBlanks = true;
  let localType: string | null = "None";
  let localRule: Record<string, unknown> = {};
  let localErrorAlert = cloneAlert(DEFAULT_ERROR);
  let localPrompt = clonePrompt(DEFAULT_PROMPT);
  let scalarLoaded = false;

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
    if (partial.touchErrorAlert) next.errorAlert = { ...localErrorAlert };
    if (partial.touchPrompt) next.prompt = { ...localPrompt };
    pending.dvWrites.set(key, next);
  }

  function notLoaded(prop: string): never {
    throw new Error(`PropertyNotLoaded:${prop}`);
  }

  const dvProxy: {
    type: string | null;
    ignoreBlanks: boolean;
    rule: Record<string, unknown>;
    errorAlert: Required<DvAlertState>;
    prompt: Required<DvPromptState>;
    load: (props?: string) => void;
    clear: () => void;
  } = {
    get type() {
      if (!scalarLoaded) notLoaded("type");
      if (options.poisonSurface && "type" in options.poisonSurface) {
        return options.poisonSurface.type as string | null;
      }
      return localType;
    },
    set type(v: string | null) {
      localType = v;
    },
    get ignoreBlanks() {
      if (!scalarLoaded) notLoaded("ignoreBlanks");
      if (options.poisonSurface && "ignoreBlanks" in options.poisonSurface) {
        return options.poisonSurface.ignoreBlanks as boolean;
      }
      return localIgnoreBlanks;
    },
    set ignoreBlanks(v: boolean) {
      options.writeCounts.ignoreBlanks += 1;
      localIgnoreBlanks = v;
      queueWrite({ ignoreBlanks: v });
    },
    get rule() {
      if (!scalarLoaded) notLoaded("rule");
      if (options.poisonSurface && "rule" in options.poisonSurface) {
        return options.poisonSurface.rule as Record<string, unknown>;
      }
      return localRule;
    },
    set rule(next: Record<string, unknown>) {
      options.writeCounts.rule += 1;
      const copy = materializeRule(next, options.keepListSourceAsRangeObject);
      const type = guessDvType(copy);
      localRule = copy;
      localType = type;
      queueWrite({ type, ignoreBlanks: localIgnoreBlanks, rule: copy });
    },
    get errorAlert() {
      if (!scalarLoaded) notLoaded("errorAlert");
      if (options.missingDvErrorAlert) {
        // Simulate missing member after "load" by throwing / undefined path
        throw new Error("PropertyNotLoaded:errorAlert");
      }
      if (options.missingDvErrorAlertFields) {
        return {} as Required<DvAlertState>;
      }
      // Plain snapshot: mutating fields must NOT queue host writes.
      return { ...localErrorAlert };
    },
    set errorAlert(v: Required<DvAlertState>) {
      options.writeCounts.errorAlert += 1;
      if (!v || typeof v !== "object") {
        throw new Error("errorAlert assignment must be an object");
      }
      localErrorAlert = cloneAlert(v);
      queueWrite({ touchErrorAlert: true });
    },
    get prompt() {
      if (!scalarLoaded) notLoaded("prompt");
      if (options.missingDvPrompt) {
        throw new Error("PropertyNotLoaded:prompt");
      }
      return { ...localPrompt };
    },
    set prompt(v: Required<DvPromptState>) {
      options.writeCounts.prompt += 1;
      if (!v || typeof v !== "object") {
        throw new Error("prompt assignment must be an object");
      }
      localPrompt = clonePrompt(v);
      queueWrite({ touchPrompt: true });
    },
    load(props?: string) {
      if (typeof props === "string") {
        options.recordLoadProps(props);
      }
      // Top-level load only (nested paths still "work" but production must not use them).
      pending.loads.push(() => {
        const committed = dvs.get(key);
        if (!committed) {
          localRule = {};
          localIgnoreBlanks = true;
          localErrorAlert = cloneAlert(DEFAULT_ERROR);
          localPrompt = clonePrompt(DEFAULT_PROMPT);
          localType = options.clearLeavesHostType ?? "None";
          scalarLoaded = true;
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
        localErrorAlert = cloneAlert(view.errorAlert ?? DEFAULT_ERROR);
        localPrompt = clonePrompt(view.prompt ?? DEFAULT_PROMPT);
        scalarLoaded = true;
      });
    },
    clear() {
      pending.dvWrites.set(key, "clear");
      localType = "None";
      localRule = {};
      localErrorAlert = cloneAlert(DEFAULT_ERROR);
      localPrompt = clonePrompt(DEFAULT_PROMPT);
      // Unload until next load+sync (matches post-mutation need to reload).
      scalarLoaded = false;
    },
  };

  // If missing members: delete properties so "in"/access fails appropriately after load
  // We use getters that throw instead.

  return dvProxy;
}
