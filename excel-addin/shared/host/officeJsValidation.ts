import type { ExcelDataValidation, ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type {
  ConditionalFormatInfo,
  ConditionalFormatRule,
  DataValidationInfo,
  DataValidationRule,
  HostResult,
} from "./types";

const CELL_VALUE = "CellValue";
const CUSTOM = "Custom";

function mapOperator(op: string): string {
  switch (op) {
    case "greaterThan":
      return "GreaterThan";
    case "lessThan":
      return "LessThan";
    case "equalTo":
      return "EqualTo";
    case "between":
      return "Between";
    case "notBetween":
      return "NotBetween";
    default:
      throw new Error(`unsupported operator: ${op}`);
  }
}

function unmapOperator(op: string | null | undefined): DataValidationRule["operator"] {
  const v = String(op ?? "").toLowerCase().replace(/\s+/g, "");
  if (v.includes("notbetween")) return "notBetween";
  if (v.includes("between")) return "between";
  if (v.includes("greater")) return "greaterThan";
  if (v.includes("less")) return "lessThan";
  if (v.includes("equal")) return "equalTo";
  return undefined;
}

function parseDvRule(dv: ExcelDataValidation): DataValidationRule | null {
  const typeStr = String(dv.type ?? "").toLowerCase();
  if (!typeStr || typeStr === "none") return null;
  const raw = dv.rule ?? {};
  if (typeStr.includes("list") || raw.list) {
    const source = raw.list?.source;
    const sourceText = typeof source === "string" ? source : String(source ?? "");
    const listValues = sourceText
      .replace(/^"|"$/g, "")
      .split(",")
      .map((s) => s.trim())
      .filter(Boolean);
    return {
      type: "list",
      formula1: sourceText || undefined,
      listValues: listValues.length > 0 ? listValues : undefined,
      allowBlank: dv.ignoreBlanks,
    };
  }
  const whole = raw.wholeNumber;
  return {
    type: "wholeNumber",
    operator: unmapOperator(whole?.operator),
    formula1: whole?.formula1 != null ? String(whole.formula1) : undefined,
    formula2: whole?.formula2 != null ? String(whole.formula2) : undefined,
    allowBlank: dv.ignoreBlanks,
  };
}

export async function officeJsListConditionalFormats(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<ConditionalFormatInfo[]>> {
  return withExcel("conditionalFormat.list", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
    range.conditionalFormats.load("items/id,items/type");
    await context.sync();
    const result: ConditionalFormatInfo[] = [];
    for (const item of range.conditionalFormats.items) {
      item.load("id,type");
      await context.sync();
      const kind = String(item.type).toLowerCase().includes("custom") ? "custom" : "cellValue";
      result.push({
        id: item.id,
        sheetName,
        range: range.address,
        kind,
        summary: `${kind}:${item.id}`,
      });
    }
    return result;
  });
}

export async function officeJsAddConditionalFormat(input: {
  sheetName: string;
  range: string;
  rule: ConditionalFormatRule;
}): Promise<HostResult<ConditionalFormatInfo>> {
  return withExcel("conditionalFormat.add", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(input.sheetName).getRange(input.range);
    range.load("address");
    const type = input.rule.kind === "custom" ? CUSTOM : CELL_VALUE;
    const cf = range.conditionalFormats.add(type);
    if (input.rule.kind === "custom") {
      if (!input.rule.formula) throw new Error("custom rule requires formula");
      if (!cf.custom) throw new Error("Custom conditional format not available");
      // Office.js: Excel.ConditionalFormatRule.formula is a string
      cf.custom.rule.formula = input.rule.formula;
      if (input.rule.fillColor) cf.custom.format.fill.color = input.rule.fillColor;
      if (input.rule.fontColor) cf.custom.format.font.color = input.rule.fontColor;
    } else {
      if (!input.rule.operator) throw new Error("cellValue rule requires operator");
      if (!input.rule.formula1) throw new Error("cellValue rule requires formula1");
      if (!cf.cellValue) throw new Error("CellValue conditional format not available");
      cf.cellValue.rule.operator = mapOperator(input.rule.operator);
      cf.cellValue.rule.formula1 = input.rule.formula1;
      if (input.rule.formula2) cf.cellValue.rule.formula2 = input.rule.formula2;
      if (input.rule.fillColor) cf.cellValue.format.fill.color = input.rule.fillColor;
      if (input.rule.fontColor) cf.cellValue.format.font.color = input.rule.fontColor;
    }
    cf.load("id,type");
    await context.sync();
    return {
      id: cf.id,
      sheetName: input.sheetName,
      range: range.address,
      kind: input.rule.kind,
      summary: `${input.rule.kind}:${cf.id}`,
    };
  });
}

export async function officeJsDeleteConditionalFormat(
  sheetName: string,
  rangeAddress: string,
  id: string,
): Promise<HostResult<{ deleted: string }>> {
  return withExcel("conditionalFormat.delete", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    const item = range.conditionalFormats.getItem(id);
    item.delete();
    await context.sync();
    return { deleted: id };
  });
}

export async function officeJsReadDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<DataValidationInfo>> {
  return withExcel("dataValidation.read", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.load("address");
    // Official surface: type / rule / ignoreBlanks (no top-level formula1/operator)
    range.dataValidation.load("type,rule,ignoreBlanks");
    await context.sync();
    return {
      sheetName,
      range: range.address,
      rule: parseDvRule(range.dataValidation),
    };
  });
}

export async function officeJsWriteDataValidation(input: {
  sheetName: string;
  range: string;
  rule: DataValidationRule;
}): Promise<HostResult<DataValidationInfo>> {
  return withExcel("dataValidation.write", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(input.sheetName).getRange(input.range);
    range.load("address");
    const dv = range.dataValidation;
    if (input.rule.type === "list") {
      const source =
        input.rule.listValues?.join(",") ??
        input.rule.formula1 ??
        (() => {
          throw new Error("list validation requires listValues or formula1");
        })();
      dv.rule = { list: { inCellDropDown: true, source } };
    } else {
      if (!input.rule.operator) throw new Error("wholeNumber requires operator");
      if (!input.rule.formula1) throw new Error("wholeNumber requires formula1");
      dv.rule = {
        wholeNumber: {
          formula1: input.rule.formula1,
          formula2: input.rule.formula2,
          operator: mapOperator(input.rule.operator),
        },
      };
    }
    dv.ignoreBlanks = input.rule.allowBlank !== false;
    dv.load("type,rule,ignoreBlanks");
    await context.sync();
    return {
      sheetName: input.sheetName,
      range: range.address,
      rule: parseDvRule(dv) ?? {
        type: input.rule.type,
        operator: input.rule.operator,
        formula1: input.rule.formula1,
        formula2: input.rule.formula2,
        listValues: input.rule.listValues,
        allowBlank: dv.ignoreBlanks,
      },
    };
  });
}

export async function officeJsClearDataValidation(
  sheetName: string,
  rangeAddress: string,
): Promise<HostResult<{ cleared: string }>> {
  return withExcel("dataValidation.clear", async (context: ExcelRequestContext) => {
    const range = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.dataValidation.clear();
    range.load("address");
    await context.sync();
    return { cleared: range.address };
  });
}
