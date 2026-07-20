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

export type DvState = {
  type: string | null;
  ignoreBlanks: boolean;
  rule: Record<string, unknown>;
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
  /** Tamper CF detail fields after commit (id/type preserved). */
  tamperCfReadback?: {
    operator?: string;
    formula1?: string;
    formula?: string;
    fillColor?: string;
    fontColor?: string;
  };
  /** Tamper DV committed state on load/readback. */
  tamperDvReadback?: {
    operator?: string;
    formula1?: string;
    allowBlank?: boolean;
    listSource?: string;
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

/**
 * Materialize DV rule for commit.
 * Range-like list.source is kept as object by default (real host readback).
 */
export function materializeRule(
  next: Record<string, unknown>,
  keepRangeObject: boolean,
): Record<string, unknown> {
  const copy: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(next)) {
    if (k === "list" && v && typeof v === "object") {
      const list = v as { inCellDropDown?: boolean; source?: unknown };
      let source = list.source;
      if (
        !keepRangeObject &&
        source &&
        typeof source === "object" &&
        source !== null &&
        "address" in source
      ) {
        source = String((source as { address: string }).address);
      }
      // Keep Range object reference as-is when keepRangeObject.
      copy.list = { inCellDropDown: list.inCellDropDown, source };
    } else if (v && typeof v === "object") {
      copy[k] = JSON.parse(JSON.stringify(v));
    } else {
      copy[k] = v;
    }
  }
  return copy;
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
  if (tamper.operator || tamper.formula1) {
    for (const key of Object.keys(next.rule)) {
      const bag = next.rule[key];
      if (bag && typeof bag === "object" && "operator" in (bag as object)) {
        const b = bag as { operator?: string; formula1?: string };
        if (tamper.operator) b.operator = tamper.operator;
        if (tamper.formula1) b.formula1 = tamper.formula1;
      }
    }
  }
  if (tamper.listSource != null && next.rule.list && typeof next.rule.list === "object") {
    (next.rule.list as { source?: unknown }).source = tamper.listSource;
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
