import {
  escapeXmlAttribute as escapeXml,
  escapeXmlTextWithQuotes as escapeXmlText,
} from "../../../shared/xmlEntities";

const OPEN_XML_FUTURE_FUNCTION_PREFIXES: Record<string, string> = {
  ANCHORARRAY: "_xlfn.ANCHORARRAY",
  BYCOL: "_xlfn.BYCOL",
  BYROW: "_xlfn.BYROW",
  CHOOSECOLS: "_xlfn.CHOOSECOLS",
  CHOOSEROWS: "_xlfn.CHOOSEROWS",
  CONCAT: "_xlfn.CONCAT",
  DROP: "_xlfn.DROP",
  EXPAND: "_xlfn.EXPAND",
  FILTER: "_xlfn._xlws.FILTER",
  HSTACK: "_xlfn.HSTACK",
  IFS: "_xlfn.IFS",
  LAMBDA: "_xlfn.LAMBDA",
  LET: "_xlfn.LET",
  MAKEARRAY: "_xlfn.MAKEARRAY",
  MAP: "_xlfn.MAP",
  MAXIFS: "_xlfn.MAXIFS",
  MINIFS: "_xlfn.MINIFS",
  RANDARRAY: "_xlfn.RANDARRAY",
  REDUCE: "_xlfn.REDUCE",
  SCAN: "_xlfn.SCAN",
  SEQUENCE: "_xlfn.SEQUENCE",
  SINGLE: "_xlfn.SINGLE",
  SORT: "_xlfn._xlws.SORT",
  SORTBY: "_xlfn.SORTBY",
  SWITCH: "_xlfn.SWITCH",
  TAKE: "_xlfn.TAKE",
  TEXTJOIN: "_xlfn.TEXTJOIN",
  TEXTSPLIT: "_xlfn.TEXTSPLIT",
  TOCOL: "_xlfn.TOCOL",
  TOROW: "_xlfn.TOROW",
  UNIQUE: "_xlfn.UNIQUE",
  VSTACK: "_xlfn.VSTACK",
  WRAPCOLS: "_xlfn.WRAPCOLS",
  WRAPROWS: "_xlfn.WRAPROWS",
  XLOOKUP: "_xlfn.XLOOKUP",
  XMATCH: "_xlfn.XMATCH",
};

const DYNAMIC_ARRAY_FUNCTIONS = new Set([
  "ANCHORARRAY",
  "BYCOL",
  "BYROW",
  "CHOOSECOLS",
  "CHOOSEROWS",
  "DROP",
  "EXPAND",
  "FILTER",
  "HSTACK",
  "MAKEARRAY",
  "MAP",
  "RANDARRAY",
  "REDUCE",
  "SCAN",
  "SEQUENCE",
  "SINGLE",
  "SORT",
  "SORTBY",
  "TAKE",
  "TEXTSPLIT",
  "TOCOL",
  "TOROW",
  "UNIQUE",
  "VSTACK",
  "WRAPCOLS",
  "WRAPROWS",
  "XLOOKUP",
  "XMATCH",
]);

export function hasDynamicArrayFormulaValue(values: unknown[][]): boolean {
  return values.some((row) => row.some((value) => (
    typeof value === "string"
    && value.startsWith("=")
    && isDynamicArrayFormula(value.slice(1))
  )));
}

export function formulaCellXml(address: string, formula: string, targetRef?: string): string {
  const normalizedFormula = normalizeFormulaForOpenXml(formula);
  if (isDynamicArrayFormula(formula)) {
    const ref = normalizeFormulaRef(targetRef, address);
    return `<c r="${address}"><f t="array" ref="${escapeXml(ref)}">${escapeXmlText(normalizedFormula)}</f></c>`;
  }
  return `<c r="${address}"><f>${escapeXmlText(normalizedFormula)}</f></c>`;
}

function normalizeFormulaForOpenXml(formula: string): string {
  let result = "";
  let token = "";
  let inString = false;

  for (let index = 0; index < formula.length; index++) {
    const char = formula[index];
    if (char === "\"") {
      result += flushFormulaToken(token);
      token = "";
      result += char;
      if (inString && formula[index + 1] === "\"") {
        result += formula[index + 1];
        index++;
      } else {
        inString = !inString;
      }
      continue;
    }
    if (!inString && /[A-Za-z0-9_.]/.test(char)) {
      token += char;
      continue;
    }
    result += flushFormulaToken(token, char);
    token = "";
    result += char;
  }

  return result + flushFormulaToken(token);
}

function flushFormulaToken(token: string, nextChar = ""): string {
  if (!token || nextChar !== "(") return token;
  if (token.includes(".")) return token;
  const prefixed = OPEN_XML_FUTURE_FUNCTION_PREFIXES[token.toUpperCase()];
  return prefixed || token;
}

function isDynamicArrayFormula(formula: string): boolean {
  return formulaFunctionNames(formula).some((name) => DYNAMIC_ARRAY_FUNCTIONS.has(name));
}

function formulaFunctionNames(formula: string): string[] {
  const names: string[] = [];
  let token = "";
  let inString = false;

  for (let index = 0; index < formula.length; index++) {
    const char = formula[index];
    if (char === "\"") {
      if (inString && formula[index + 1] === "\"") {
        index++;
      } else {
        inString = !inString;
      }
      token = "";
      continue;
    }
    if (!inString && /[A-Za-z0-9_.]/.test(char)) {
      token += char;
      continue;
    }
    if (!inString && char === "(" && token) {
      const bareName = token.split(".").pop() || token;
      names.push(bareName.toUpperCase());
    }
    token = "";
  }

  return names;
}

function normalizeFormulaRef(targetRef: string | undefined, address: string): string {
  const ref = (targetRef || "").trim();
  if (/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(ref)) return ref.toUpperCase();
  return address;
}
