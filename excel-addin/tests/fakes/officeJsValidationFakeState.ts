/** Shared CF/DV fake state types and pure helpers. */

export type CfState = {
  id: string;
  type: string;
  cellRule?: { formula1: string; formula2?: string; operator: string };
  cellFill?: string;
  cellFont?: string;
  customFormula?: string;
  customFill?: string;
  customFont?: string;
};

export type DvAlertState = {
  showAlert?: boolean;
  style?: string;
  title?: string;
  message?: string;
};

export type DvPromptState = {
  showPrompt?: boolean;
  title?: string;
  message?: string;
};

export type DvState = {
  type: string | null;
  ignoreBlanks: boolean;
  rule: Record<string, unknown>;
  errorAlert?: DvAlertState;
  prompt?: DvPromptState;
};

export type CtxPending = {
  cfAdds: Array<{ key: string; state: CfState }>;
  cfDeletes: Array<{ key: string; id: string }>;
  cfPatches: Array<{ key: string; id: string; patch: Partial<CfState> }>;
  dvWrites: Map<string, DvState | "clear">;
  loads: Array<() => void>;
};

export type ValidationFakeOptions = {
  excelApi16?: boolean;
  excelApi18?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  failAddSync?: boolean;
  failDeleteReadback?: boolean;
  failWriteSync?: boolean;
  failClearReadback?: boolean;
  /** After clear, leave hostType as this instead of None (e.g. Inconsistent). */
  clearLeavesHostType?: string;
  seedContainsText?: boolean;
  seedInconsistentDv?: boolean;
  seedManyCf?: number;
  /** Keep list.source as Range-like object after commit (default true). */
  keepListSourceAsRangeObject?: boolean;
  /** Omit dataValidation.errorAlert member (precheck fail). */
  missingDvErrorAlert?: boolean;
  /** Omit dataValidation.prompt member (precheck fail). */
  missingDvPrompt?: boolean;
  /** Omit nested errorAlert.style etc. */
  missingDvErrorAlertFields?: boolean;
  /** Tamper CF detail fields after commit (id/type preserved). */
  tamperCfReadback?: {
    operator?: string;
    formula1?: string;
    formula2?: string;
    formula?: string;
    fillColor?: string;
    fontColor?: string;
  };
  /** Tamper DV committed state on load/readback. */
  tamperDvReadback?: {
    operator?: string;
    formula1?: string;
    formula2?: string;
    allowBlank?: boolean;
    listSource?: unknown;
    errorAlertStyle?: string;
    errorAlertTitle?: string;
    errorAlertMessage?: string;
    errorAlertShow?: boolean;
    promptTitle?: string;
    promptMessage?: string;
    promptShow?: boolean;
  };
};

export function keyOf(sheet: string, address: string): string {
  return `${sheet}!${address.toUpperCase().replace(/\$/g, "")}`;
}

export function guessDvType(rule: Record<string, unknown>): string {
  if (rule.list) return "List";
  if (rule.wholeNumber) return "WholeNumber";
  if (rule.decimal) return "Decimal";
  if (rule.date) return "Date";
  if (rule.time) return "Time";
  if (rule.textLength) return "TextLength";
  if (rule.custom) return "Custom";
  return "None";
}

export type RangeAddressToken = { __rangeAddress: string };

/**
 * Materialize DV rule for commit.
 * ClientObject Range → { __rangeAddress } token (never read .address here).
 * On load, token is rehydrated to a load-gated ClientObject for the current context.
 */
export function materializeRule(
  next: Record<string, unknown>,
  keepRangeObject: boolean,
): Record<string, unknown> {
  const copy: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(next)) {
    if (k === "list" && v && typeof v === "object") {
      const list = v as { inCellDropDown?: boolean; source?: unknown };
      let source: unknown = list.source;
      if (source && typeof source === "object" && source !== null) {
        const token = (source as { __rangeAddress?: string }).__rangeAddress;
        if (typeof token === "string") {
          source = keepRangeObject ? { __rangeAddress: token } : token;
        } else if (typeof (source as { load?: unknown }).load === "function") {
          // ClientObject: freeze via stamped __rangeAddress only (do not read .address).
          const stamped = (source as { __rangeAddress?: string }).__rangeAddress;
          if (typeof stamped !== "string") {
            throw new Error("list Range source missing __rangeAddress stamp");
          }
          source = keepRangeObject ? { __rangeAddress: stamped } : stamped;
        }
      }
      copy.list = { inCellDropDown: list.inCellDropDown, source };
    } else if (v && typeof v === "object") {
      copy[k] = JSON.parse(JSON.stringify(v));
    } else {
      copy[k] = v;
    }
  }
  return copy;
}

/** Build load-gated ClientObject Range for current context pending. */
export function makeLoadGatedRange(
  committedAddress: string,
  pending: { loads: Array<() => void> },
): { __rangeAddress: string; load: (props: string) => void; address: string } {
  let loaded = false;
  return {
    __rangeAddress: committedAddress,
    load(props: string) {
      if (props.includes("address")) {
        pending.loads.push(() => {
          loaded = true;
        });
      }
    },
    get address() {
      if (!loaded) throw new Error("PropertyNotLoaded:address");
      return committedAddress;
    },
  };
}

export function rehydrateDvRule(
  rule: Record<string, unknown>,
  pending: { loads: Array<() => void> },
  keepRangeObject: boolean,
): Record<string, unknown> {
  const list = rule.list;
  if (!keepRangeObject || !list || typeof list !== "object") {
    return rule;
  }
  const src = (list as { source?: unknown }).source;
  if (src && typeof src === "object" && src !== null) {
    const token = (src as { __rangeAddress?: string }).__rangeAddress;
    if (typeof token === "string") {
      return {
        ...rule,
        list: {
          ...(list as object),
          source: makeLoadGatedRange(token, pending),
        },
      };
    }
  }
  if (typeof src === "string" && /!/.test(src)) {
    // address string committed while keepRangeObject — still rehydrate as ClientObject
    return {
      ...rule,
      list: {
        ...(list as object),
        source: makeLoadGatedRange(src, pending),
      },
    };
  }
  return rule;
}

export function seedContainsText(cfs: Map<string, CfState[]>): void {
  cfs.set(keyOf("Sheet1", "A1:A10"), [
    { id: "cf_contains", type: "ContainsText" },
    {
      id: "cf_cell",
      type: "CellValue",
      cellRule: { operator: "GreaterThan", formula1: "0" },
      cellFill: "#FFFFFF",
      cellFont: "#000000",
    },
  ]);
}

export function seedManyCf(cfs: Map<string, CfState[]>, count: number): void {
  const many: CfState[] = [];
  for (let i = 0; i < count; i += 1) {
    many.push({
      id: `cf_many_${i}`,
      type: i % 2 === 0 ? "CellValue" : "DataBar",
      cellRule:
        i % 2 === 0 ? { operator: "EqualTo", formula1: String(i) } : undefined,
    });
  }
  cfs.set(keyOf("Sheet1", "A1:A100"), many);
}

export function seedInconsistentDv(dvs: Map<string, DvState>): void {
  dvs.set(keyOf("Sheet1", "B1:B5"), {
    type: "Inconsistent",
    ignoreBlanks: true,
    rule: {},
  });
}

export function applyDvTamper(
  state: DvState,
  tamper: ValidationFakeOptions["tamperDvReadback"],
): DvState {
  if (!tamper) return state;
  // Shallow-clone rule bags but keep Range-like list.source by reference.
  const rule: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(state.rule)) {
    if (k === "list" && v && typeof v === "object") {
      const list = v as { inCellDropDown?: boolean; source?: unknown };
      rule.list = { inCellDropDown: list.inCellDropDown, source: list.source };
    } else if (v && typeof v === "object") {
      rule[k] = { ...(v as object) };
    } else {
      rule[k] = v;
    }
  }
  const next: DvState = {
    type: state.type,
    ignoreBlanks:
      tamper.allowBlank !== undefined ? tamper.allowBlank : state.ignoreBlanks,
    rule,
  };
  if (tamper.operator || tamper.formula1 || tamper.formula2 != null) {
    for (const key of Object.keys(next.rule)) {
      const bag = next.rule[key];
      if (!bag || typeof bag !== "object") continue;
      const b = bag as {
        operator?: string;
        formula1?: string;
        formula?: string;
        formula2?: string;
      };
      if (tamper.operator && "operator" in b) b.operator = tamper.operator;
      if (tamper.formula1) {
        if ("formula1" in b) b.formula1 = tamper.formula1;
        if ("formula" in b && key === "custom") b.formula = tamper.formula1;
      }
      if (tamper.formula2 != null) {
        // Inject on compare bags and custom (host may surface formula2 illegally).
        b.formula2 = tamper.formula2;
      }
    }
  }
  if (tamper.listSource != null && next.rule.list && typeof next.rule.list === "object") {
    (next.rule.list as { source?: unknown }).source = tamper.listSource;
  }
  if (
    tamper.errorAlertStyle != null ||
    tamper.errorAlertTitle != null ||
    tamper.errorAlertMessage != null ||
    tamper.errorAlertShow !== undefined
  ) {
    next.errorAlert = { ...(state.errorAlert ?? {}) };
    if (tamper.errorAlertStyle != null) next.errorAlert.style = tamper.errorAlertStyle;
    if (tamper.errorAlertTitle != null) next.errorAlert.title = tamper.errorAlertTitle;
    if (tamper.errorAlertMessage != null) next.errorAlert.message = tamper.errorAlertMessage;
    if (tamper.errorAlertShow !== undefined) next.errorAlert.showAlert = tamper.errorAlertShow;
  }
  if (
    tamper.promptTitle != null ||
    tamper.promptMessage != null ||
    tamper.promptShow !== undefined
  ) {
    next.prompt = { ...(state.prompt ?? {}) };
    if (tamper.promptTitle != null) next.prompt.title = tamper.promptTitle;
    if (tamper.promptMessage != null) next.prompt.message = tamper.promptMessage;
    if (tamper.promptShow !== undefined) next.prompt.showPrompt = tamper.promptShow;
  }
  return next;
}

export function applyCfTamper(
  state: CfState,
  tamper: ValidationFakeOptions["tamperCfReadback"],
): CfState {
  if (!tamper) return state;
  const next = { ...state };
  if (next.cellRule) {
    next.cellRule = { ...next.cellRule };
    if (tamper.operator) next.cellRule.operator = tamper.operator;
    if (tamper.formula1) next.cellRule.formula1 = tamper.formula1;
    if (tamper.formula2 != null) next.cellRule.formula2 = tamper.formula2;
  }
  if (tamper.formula != null) next.customFormula = tamper.formula;
  if (tamper.fillColor != null) {
    next.cellFill = tamper.fillColor;
    next.customFill = tamper.fillColor;
  }
  if (tamper.fontColor != null) {
    next.cellFont = tamper.fontColor;
    next.customFont = tamper.fontColor;
  }
  return next;
}
