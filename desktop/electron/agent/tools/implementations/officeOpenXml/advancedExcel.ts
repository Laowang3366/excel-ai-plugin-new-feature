/**
 * Excel Open XML 高级操作
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: 将统一 Office action 路由到本模块。
 * - tableStyler.ts: 复用既有表格样式能力，避免重复实现表格美化。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import { parseOfficeLocator } from "../../officeCore/locator";
import { doneResult, failedResult, needsComResult, unsupportedResult } from "../../officeCore/results";
import type { OfficeActionKind, OfficeActionResult } from "../../officeCore/types";
import {
  decodeXmlText as decodeXml,
  escapeXmlAttribute as escapeXml,
  escapeXmlTextWithQuotes as escapeXmlText,
} from "../../../shared/xmlEntities";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";
import type { OfficeOpenXmlTableStylePreset } from "./types";

export interface ExcelAdvancedActionInput {
  operation: string;
  filePath: string;
  outputPath?: string;
  target?: string;
  action?: OfficeActionKind;
  params?: Record<string, unknown>;
}

const TABLE_STYLES = new Set<OfficeOpenXmlTableStylePreset>(["professional", "compact", "financial"]);
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

export async function applyExcelAdvancedAction(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  try {
    if (input.operation === "createWorkbook") {
      return await createWorkbook(input);
    }
    if (input.operation === "writeRange") {
      return await writeRange(input);
    }
    if (input.operation === "setDataValidation") {
      return await setDataValidation(input);
    }
    if (input.operation === "applyConditionalFormatting") {
      return await applyConditionalFormatting(input);
    }
    if (input.operation === "styleTable") {
      return await styleTable(input);
    }
    if (input.operation === "insertChart") {
      return needsComResult({
        app: "excel",
        action: input.action || "insert",
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: "Open XML 图表包生成尚未覆盖，需要显式 COM 兜底",
      });
    }

    return unsupportedResult({
      app: "excel",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: `暂不支持 Excel Open XML 高级操作: ${input.operation}`,
    });
  } catch (error) {
    return failedResult({
      app: "excel",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
    }, error);
  }
}

async function createWorkbook(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  const outputPath = input.outputPath || input.filePath;
  const sheetNames = normalizeSheetNames(input.params?.sheetNames);
  const values = normalizeValues(input.params?.values);
  const containsDynamicFormula = hasDynamicArrayFormulaValue(values);
  const target = input.target || `range:${sheetNames[0]}!${stringParam(input.params, "startCell") || "A1"}`;
  const locator = parseOfficeLocator(target);
  const targetSheetName = locator?.sheetName || sheetNames[0];
  const startCell = firstCellAddress(locator?.address || "A1");

  const zip = new JSZip();
  addWorkbookBaseParts(zip, sheetNames);
  for (let index = 0; index < sheetNames.length; index++) {
    const sheetName = sheetNames[index];
    const sheetData = sheetName === targetSheetName && values.length > 0
      ? buildSheetDataXml(startCell, values, locator?.address || startCell, containsDynamicFormula)
      : "";
    zip.file(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheetData));
  }

  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return excelDone(input, outputPath, [
    "xl/workbook.xml",
    ...sheetNames.map((_name, index) => `xl/worksheets/sheet${index + 1}.xml`),
  ], "已使用内置 Open XML 创建 Excel 工作簿", {
    engine: "openxml",
    operation: "createWorkbook",
    filePath: input.filePath,
    outputPath,
    sheetNames,
    initialRange: values.length > 0 ? target : undefined,
    rowsWritten: values.length,
  });
}

async function writeRange(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  const values = normalizeValues(input.params?.values);
  if (values.length === 0) {
    return failedResult({
      app: "excel",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
    }, "writeRange 操作需要 params.values 二维数组");
  }

  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const containsDynamicFormula = hasDynamicArrayFormulaValue(values);
  const partName = await resolveWorksheetPart(zip, input.target);
  const part = zip.file(partName);
  if (!part) throw new Error(`找不到工作表部件: ${partName}`);

  const startCell = targetStartCell(input.target);
  const xml = await part.async("text");
  zip.file(partName, mergeSheetDataXml(xml, startCell, values, targetAddress(input.target), containsDynamicFormula));

  const outputPath = input.outputPath || input.filePath;
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return excelDone(input, outputPath, [partName], "已使用内置 Open XML 写入 Excel 单元格", {
    engine: "openxml",
    operation: "writeRange",
    filePath: input.filePath,
    outputPath,
    target: input.target,
    rowsWritten: values.length,
    columnsWritten: Math.max(...values.map((row) => row.length)),
  });
}

async function setDataValidation(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const partName = await resolveWorksheetPart(zip, input.target);
  const part = zip.file(partName);
  if (!part) throw new Error(`找不到工作表部件: ${partName}`);

  const address = targetAddress(input.target);
  const values = Array.isArray(input.params?.values)
    ? input.params.values.map((item) => String(item)).join(",")
    : "";
  const type = typeof input.params?.type === "string" ? input.params.type : "list";
  const xml = await part.async("text");
  const validationXml = `<dataValidations count="1"><dataValidation type="${escapeXml(type)}" allowBlank="1" sqref="${escapeXml(address)}"><formula1>"${escapeXml(values)}"</formula1></dataValidation></dataValidations>`;
  zip.file(partName, insertBeforeWorksheetEnd(removeExistingDataValidations(xml), validationXml));

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return excelDone(input, outputPath, [partName], "已写入 Excel 数据验证");
}

async function applyConditionalFormatting(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const partName = await resolveWorksheetPart(zip, input.target);
  const part = zip.file(partName);
  if (!part) throw new Error(`找不到工作表部件: ${partName}`);

  const address = targetAddress(input.target);
  const color = typeof input.params?.fillColor === "string" ? input.params.fillColor : "FFF2CC";
  const xml = await part.async("text");
  const formattingXml = `<conditionalFormatting sqref="${escapeXml(address)}"><cfRule type="expression" priority="1"><formula>TRUE</formula><dxf><fill><patternFill patternType="solid"><fgColor rgb="FF${escapeXml(color)}" /></patternFill></fill></dxf></cfRule></conditionalFormatting>`;
  zip.file(partName, insertBeforeWorksheetEnd(xml, formattingXml));

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return excelDone(input, outputPath, [partName], "已写入 Excel 条件格式");
}

async function styleTable(input: ExcelAdvancedActionInput): Promise<OfficeActionResult> {
  const style = normalizeTableStyle(input.params?.style);
  const result = await applyOfficeOpenXmlTableStyle({
    filePath: input.filePath,
    outputPath: input.outputPath,
    target: input.target,
    style,
  });
  return excelDone(input, result.outputPath, result.changedParts, "已应用 Excel 表格样式", result);
}

function excelDone(
  input: ExcelAdvancedActionInput,
  outputPath: string,
  changedParts: string[],
  summary: string,
  data?: unknown
): OfficeActionResult {
  return doneResult({
    engine: "openxml",
    app: "excel",
    action: input.action || "edit",
    operation: input.operation,
    filePath: input.filePath,
    outputPath,
    target: input.target,
    summary,
    data,
    validation: {
      ok: true,
      checks: [{ name: "output-file", ok: true, message: "已生成输出文件" }],
    },
    changes: changedParts.map((partName) => ({
      kind: "openxml-part",
      target: partName,
      detail: `已更新 ${partName}`,
    })),
  });
}

async function resolveWorksheetPart(zip: JSZip, target?: string): Promise<string> {
  const locator = target ? parseOfficeLocator(target) : undefined;
  if (locator?.kind === "range" && locator.sheetName) {
    const workbookPart = zip.file("xl/workbook.xml");
    const relsPart = zip.file("xl/_rels/workbook.xml.rels");
    if (!workbookPart || !relsPart) return "xl/worksheets/sheet1.xml";

    const workbookXml = await workbookPart.async("text");
    const relsXml = await relsPart.async("text");
    const relId = findSheetRelationshipId(workbookXml, locator.sheetName);
    if (!relId) throw new Error(`找不到工作表: ${locator.sheetName}`);

    const targetPath = findRelationshipTarget(relsXml, relId);
    if (!targetPath) throw new Error(`找不到工作表关系: ${relId}`);
    return normalizeWorkbookTarget(targetPath);
  }
  return "xl/worksheets/sheet1.xml";
}

function findSheetRelationshipId(workbookXml: string, sheetName: string): string | undefined {
  const sheetRe = /<sheet\b[^>]*>/g;
  let match: RegExpExecArray | null;
  while ((match = sheetRe.exec(workbookXml))) {
    const attrs = parseXmlAttributes(match[0]);
    if (decodeXml(attrs.name || "") === sheetName) {
      return attrs["r:id"] || attrs.id;
    }
  }
  return undefined;
}

function findRelationshipTarget(relsXml: string, relId: string): string | undefined {
  const relRe = /<Relationship\b[^>]*>/g;
  let match: RegExpExecArray | null;
  while ((match = relRe.exec(relsXml))) {
    const attrs = parseXmlAttributes(match[0]);
    if (attrs.Id === relId) return attrs.Target;
  }
  return undefined;
}

function parseXmlAttributes(tagXml: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  const attrRe = /([\w:-]+)\s*=\s*["']([^"']*)["']/g;
  let match: RegExpExecArray | null;
  while ((match = attrRe.exec(tagXml))) {
    attrs[match[1]] = match[2];
  }
  return attrs;
}

function normalizeWorkbookTarget(targetPath: string): string {
  const normalized = targetPath.replace(/\\/g, "/").replace(/^\/+/, "");
  if (normalized.startsWith("xl/")) return normalized;
  return path.posix.join("xl", normalized);
}

function targetAddress(target?: string): string {
  const locator = target ? parseOfficeLocator(target) : undefined;
  return locator?.address || "A1";
}

function targetStartCell(target?: string): string {
  return firstCellAddress(targetAddress(target));
}

function firstCellAddress(address: string): string {
  return address.split(":")[0]?.trim() || "A1";
}

function removeExistingDataValidations(xml: string): string {
  return xml.replace(/<dataValidations\b[\s\S]*?<\/dataValidations>/g, "");
}

function insertBeforeWorksheetEnd(xml: string, addition: string): string {
  return xml.includes("</worksheet>") ? xml.replace("</worksheet>", `${addition}</worksheet>`) : `${xml}${addition}`;
}

function normalizeSheetNames(value: unknown): string[] {
  if (!Array.isArray(value)) return ["Sheet1"];
  const names = value
    .filter((item): item is string => typeof item === "string" && item.trim().length > 0)
    .map((name) => name.trim());
  return names.length > 0 ? names : ["Sheet1"];
}

function normalizeValues(value: unknown): unknown[][] {
  if (!Array.isArray(value)) return [];
  if (value.every((row) => Array.isArray(row))) return value as unknown[][];
  return [value];
}

function hasDynamicArrayFormulaValue(values: unknown[][]): boolean {
  return values.some((row) => row.some((value) => (
    typeof value === "string"
    && value.startsWith("=")
    && isDynamicArrayFormula(value.slice(1))
  )));
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
}

function addWorkbookBaseParts(zip: JSZip, sheetNames: string[]): void {
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${sheetNames.map((_name, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("\n  ")}
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetNames.map((name, index) => `<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("\n    ")}
  </sheets>
</workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetNames.map((_name, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("\n  ")}
  <Relationship Id="rId${sheetNames.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`);
}

function worksheetXml(sheetData: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetData}</sheetData>
</worksheet>`;
}

function buildSheetDataXml(
  startCell: string,
  values: unknown[][],
  targetRef?: string,
  clearBlankCells = false
): string {
  const start = parseCellAddress(startCell);
  return values.map((row, rowOffset) => {
    const rowNumber = start.row + rowOffset;
    const cells = row
      .map((value, colOffset) => cellXml(toCellAddress(start.col + colOffset, rowNumber), value, targetRef, clearBlankCells))
      .filter(Boolean)
      .join("");
    return `<row r="${rowNumber}">${cells}</row>`;
  }).join("");
}

function replaceSheetData(xml: string, sheetData: string): string {
  if (/<sheetData\b[\s\S]*?<\/sheetData>/.test(xml)) {
    return xml.replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/, `<sheetData>${sheetData}</sheetData>`);
  }
  return insertBeforeWorksheetEnd(xml, `<sheetData>${sheetData}</sheetData>`);
}

function mergeSheetDataXml(
  xml: string,
  startCell: string,
  values: unknown[][],
  targetRef?: string,
  clearBlankCells = false
): string {
  const cellsByRow = new Map<number, Map<number, string>>();
  const sheetDataMatch = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
  if (sheetDataMatch) {
    const rowRe = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
    let rowMatch: RegExpExecArray | null;
    while ((rowMatch = rowRe.exec(sheetDataMatch[1]))) {
      const rowNumber = Number(rowMatch[1]);
      const rowCells = getOrCreateRow(cellsByRow, rowNumber);
      const cellRe = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:>[\s\S]*?<\/c>|\/>)/gi;
      let cellMatch: RegExpExecArray | null;
      while ((cellMatch = cellRe.exec(rowMatch[2]))) {
        const parsed = parseCellAddress(cellMatch[1]);
        rowCells.set(parsed.col, cellMatch[0]);
      }
    }
  }

  const start = parseCellAddress(startCell);
  values.forEach((row, rowOffset) => {
    const rowNumber = start.row + rowOffset;
    const rowCells = getOrCreateRow(cellsByRow, rowNumber);
    row.forEach((value, colOffset) => {
      const colNumber = start.col + colOffset;
      const nextCellXml = cellXml(toCellAddress(colNumber, rowNumber), value, targetRef, clearBlankCells);
      if (nextCellXml) {
        rowCells.set(colNumber, nextCellXml);
      } else {
        rowCells.delete(colNumber);
      }
    });
  });

  const mergedSheetData = [...cellsByRow.entries()]
    .sort(([a], [b]) => a - b)
    .map(([rowNumber, cells]) => {
      const cellXmls = [...cells.entries()]
        .sort(([a], [b]) => a - b)
        .map(([, cell]) => cell)
        .join("");
      return `<row r="${rowNumber}">${cellXmls}</row>`;
    })
    .join("");
  return replaceSheetData(xml, mergedSheetData);
}

function getOrCreateRow(rows: Map<number, Map<number, string>>, rowNumber: number): Map<number, string> {
  const existing = rows.get(rowNumber);
  if (existing) return existing;
  const next = new Map<number, string>();
  rows.set(rowNumber, next);
  return next;
}

function cellXml(address: string, value: unknown, targetRef?: string, clearBlankCells = false): string {
  if (value === null || value === undefined) return clearBlankCells ? "" : `<c r="${address}"/>`;
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${address}"><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${address}" t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  const text = String(value);
  if (clearBlankCells && text === "") return "";
  if (text.startsWith("=")) {
    return formulaCellXml(address, text.slice(1), targetRef);
  }
  return `<c r="${address}" t="inlineStr"><is><t>${escapeXmlText(text)}</t></is></c>`;
}

function formulaCellXml(address: string, formula: string, targetRef?: string): string {
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

function parseCellAddress(address: string): { col: number; row: number } {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) return { col: 1, row: 1 };
  return { col: columnNameToNumber(match[1]), row: Number(match[2]) };
}

function columnNameToNumber(name: string): number {
  return name.toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
}

function toCellAddress(col: number, row: number): string {
  let value = col;
  let name = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    value = Math.floor((value - 1) / 26);
  }
  return `${name}${row}`;
}

function normalizeTableStyle(value: unknown): OfficeOpenXmlTableStylePreset {
  return typeof value === "string" && TABLE_STYLES.has(value as OfficeOpenXmlTableStylePreset)
    ? value as OfficeOpenXmlTableStylePreset
    : "professional";
}

function defaultOutputPath(filePath: string): string {
  const ext = path.extname(filePath);
  return path.join(path.dirname(filePath), `${path.basename(filePath, ext)}-advanced${ext}`);
}
