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

/** Materialize DV rule for commit; Range-like list.source becomes address string. */
export function materializeRule(next: Record<string, unknown>): Record<string, unknown> {
  const copy: Record<string, unknown> = {};
  for (const [k, v] of Object.entries(next)) {
    if (k === "list" && v && typeof v === "object") {
      const list = v as { inCellDropDown?: boolean; source?: unknown };
      let source = list.source;
      if (source && typeof source === "object" && source !== null && "address" in source) {
        source = String((source as { address: string }).address);
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
