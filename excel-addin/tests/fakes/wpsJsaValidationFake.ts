/**
 * Minimal WPS JSA FormatConditions + Validation fake.
 * COM-style indexes are 1-based; Validation.Type uses xlValidate* constants.
 */
import {
  XL_VALIDATE_INPUT_ONLY,
  XL_VALIDATE_LIST,
} from "../../shared/host/wpsJsaValidationConstants";

export type WpsValidationFakeOptions = {
  missingFormatConditions?: boolean;
  missingValidation?: boolean;
  missingFcAdd?: boolean;
  missingFcDelete?: boolean;
  missingFcItem?: boolean;
  missingValidationAdd?: boolean;
  missingValidationDelete?: boolean;
  tamperCf?: {
    operator?: number;
    formula1?: string;
    formula2?: string;
    fillColor?: number;
    fontColor?: number;
  };
  tamperDv?: {
    type?: number;
    operator?: number;
    formula1?: string;
    formula2?: string;
    ignoreBlank?: boolean;
  };
  /** Add throws after Delete (restore path). */
  addThrows?: string;
  /** After clear Delete, leave residual type instead of None. */
  clearLeavesType?: number;
  validationDeleteThrows?: string;
};

type CondRec = {
  Type: number;
  Operator?: number;
  Formula1?: string;
  Formula2?: string;
  fillOle?: number;
  fontOle?: number;
};

type ValRec = {
  Type: number;
  Operator?: number;
  Formula1?: string;
  Formula2?: string;
  IgnoreBlank: boolean;
  InCellDropdown: boolean;
} | null;

export function installWpsValidationFake(options: WpsValidationFakeOptions = {}) {
  const conditions = new Map<string, CondRec[]>();
  const validations = new Map<string, ValRec>();
  let addThrowRemaining = options.addThrows ? 1 : 0;
  const keyOf = (sheet: string, address: string) => {
    const bare = String(address).includes("!")
      ? String(address).split("!").pop()!
      : String(address);
    return `${sheet}!${bare}`;
  };

  function rangeProxy(sheetName: string, address: string) {
    const bare = String(address).includes("!")
      ? String(address).split("!").pop()!
      : String(address);
    const key = keyOf(sheetName, bare);
    const proxy: Record<string, unknown> = {
      Address: `${sheetName}!${bare}`,
    };

    if (!options.missingFormatConditions) {
      const fc: Record<string, unknown> = {
        get Count() {
          return (conditions.get(key) ?? []).length;
        },
      };
      if (!options.missingFcItem) {
        fc.Item = (index: number) => {
          const list = conditions.get(key) ?? [];
          const rec = list[index - 1];
          if (!rec) throw new Error(`FormatConditions.Item(${index}) missing`);
          return conditionProxy(rec, () => {
            const arr = conditions.get(key) ?? [];
            const i = arr.indexOf(rec);
            if (i >= 0) arr.splice(i, 1);
            conditions.set(key, arr);
          });
        };
      }
      if (!options.missingFcAdd) {
        fc.Add = (
          type: number,
          operator?: number,
          formula1?: string,
          formula2?: string,
        ) => {
          const list = conditions.get(key) ?? [];
          const rec: CondRec = {
            Type: type,
            Operator: operator,
            Formula1: formula1 != null ? String(formula1) : undefined,
            Formula2: formula2 != null ? String(formula2) : undefined,
          };
          if (options.tamperCf) {
            if (options.tamperCf.operator != null) rec.Operator = options.tamperCf.operator;
            if (options.tamperCf.formula1 != null) rec.Formula1 = options.tamperCf.formula1;
            if (options.tamperCf.formula2 != null) rec.Formula2 = options.tamperCf.formula2;
            if (options.tamperCf.fillColor != null) rec.fillOle = options.tamperCf.fillColor;
            if (options.tamperCf.fontColor != null) rec.fontOle = options.tamperCf.fontColor;
          }
          list.push(rec);
          conditions.set(key, list);
          return conditionProxy(rec, () => {
            const arr = conditions.get(key) ?? [];
            const i = arr.indexOf(rec);
            if (i >= 0) arr.splice(i, 1);
            conditions.set(key, arr);
          });
        };
      }
      if (!options.missingFcDelete) {
        fc.Delete = () => {
          conditions.set(key, []);
        };
      }
      proxy.FormatConditions = fc;
    }

    if (!options.missingValidation) {
      proxy.Validation = validationProxy(key);
    }
    return proxy;
  }

  function conditionProxy(rec: CondRec, onDelete: () => void) {
    const p: Record<string, unknown> = {
      get Type() {
        return rec.Type;
      },
      get Operator() {
        return rec.Operator;
      },
      get Formula1() {
        return rec.Formula1;
      },
      get Formula2() {
        return rec.Formula2;
      },
      Interior: {
        get Color(): number {
          if (options.tamperCf?.fillColor != null) return options.tamperCf.fillColor;
          return rec.fillOle ?? 0;
        },
        set Color(v: number) {
          rec.fillOle = v;
        },
      },
      Font: {
        get Color(): number {
          if (options.tamperCf?.fontColor != null) return options.tamperCf.fontColor;
          return rec.fontOle ?? 0;
        },
        set Color(v: number) {
          rec.fontOle = v;
        },
      },
    };
    if (!options.missingFcDelete) {
      p.Delete = () => onDelete();
    }
    return p;
  }

  function validationProxy(key: string) {
    if (!validations.has(key)) validations.set(key, null);

    const p: Record<string, unknown> = {
      get Type() {
        const r = validations.get(key) ?? null;
        if (!r) return XL_VALIDATE_INPUT_ONLY;
        return r.Type;
      },
      get Operator() {
        return (validations.get(key) ?? null)?.Operator;
      },
      get Formula1() {
        return (validations.get(key) ?? null)?.Formula1 ?? "";
      },
      get Formula2() {
        return (validations.get(key) ?? null)?.Formula2 ?? "";
      },
      get IgnoreBlank() {
        return (validations.get(key) ?? null)?.IgnoreBlank ?? true;
      },
      set IgnoreBlank(v: boolean) {
        const r = validations.get(key);
        if (r) r.IgnoreBlank = v;
      },
      get InCellDropdown() {
        return (validations.get(key) ?? null)?.InCellDropdown ?? true;
      },
      set InCellDropdown(v: boolean) {
        const r = validations.get(key);
        if (r) r.InCellDropdown = v;
      },
    };

    if (!options.missingValidationDelete) {
      p.Delete = () => {
        if (options.validationDeleteThrows) {
          throw new Error(options.validationDeleteThrows);
        }
        if (options.clearLeavesType != null) {
          validations.set(key, {
            Type: options.clearLeavesType,
            Formula1: "residual",
            IgnoreBlank: true,
            InCellDropdown: true,
          });
          return;
        }
        validations.set(key, null);
      };
    }
    if (!options.missingValidationAdd) {
      p.Add = (
        type: number,
        _alert?: number,
        operator?: number,
        formula1?: string,
        formula2?: string,
      ) => {
        if (addThrowRemaining > 0) {
          addThrowRemaining -= 1;
          throw new Error(options.addThrows!);
        }
        const rec: NonNullable<ValRec> = {
          Type: type,
          Operator: operator,
          Formula1: formula1 != null ? String(formula1) : undefined,
          Formula2: formula2 != null ? String(formula2) : undefined,
          IgnoreBlank: true,
          InCellDropdown: type === XL_VALIDATE_LIST,
        };
        if (options.tamperDv) {
          if (options.tamperDv.type != null) rec.Type = options.tamperDv.type;
          if (options.tamperDv.operator != null) rec.Operator = options.tamperDv.operator;
          if (options.tamperDv.formula1 != null) rec.Formula1 = options.tamperDv.formula1;
          if (options.tamperDv.formula2 != null) rec.Formula2 = options.tamperDv.formula2;
          if (options.tamperDv.ignoreBlank != null) {
            rec.IgnoreBlank = options.tamperDv.ignoreBlank;
          }
        }
        validations.set(key, rec);
      };
    }
    return p;
  }

  const sheet = {
    Name: "Sheet1",
    Index: 1,
    Range: (address: string) => rangeProxy("Sheet1", address),
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Application: unknown }).Application = {
    Name: "WPS 表格",
    ActiveWorkbook: {
      Name: "Book1",
      ActiveSheet: sheet,
      Worksheets: {
        Count: 1,
        Item: (nameOrIndex: string | number) => {
          if (nameOrIndex === 1 || nameOrIndex === "Sheet1") return sheet;
          throw new Error("sheet not found");
        },
      },
    },
  };

  return {
    seedCondition(sheetName: string, address: string, rec: CondRec) {
      const key = keyOf(sheetName, address);
      const list = conditions.get(key) ?? [];
      list.push(rec);
      conditions.set(key, list);
    },
    seedValidation(sheetName: string, address: string, rec: NonNullable<ValRec>) {
      validations.set(keyOf(sheetName, address), { ...rec });
    },
    getConditions(sheetName: string, address: string) {
      return [...(conditions.get(keyOf(sheetName, address)) ?? [])];
    },
    getValidation(sheetName: string, address: string) {
      return validations.get(keyOf(sheetName, address)) ?? null;
    },
  };
}

export function uninstallWpsValidationFake() {
  delete (globalThis as { Application?: unknown }).Application;
}
