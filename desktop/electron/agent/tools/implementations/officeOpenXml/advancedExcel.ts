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
import { failedResult, needsComResult, unsupportedResult } from "../../officeCore/results";
import type { OfficeActionKind, OfficeActionResult } from "../../officeCore/types";
import {
  decodeXmlText as decodeXml,
  escapeXmlAttribute as escapeXml,
  parseXmlAttributes,
} from "../../../shared/xmlEntities";
import { hasDynamicArrayFormulaValue } from "./excelFormulaXml";
import {
  addWorkbookBaseParts,
  buildSheetDataXml,
  mergeSheetDataXml,
  worksheetXml,
} from "./excelSheetXml";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";
import { createOpenXmlDoneResult } from "./actionResult";
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
const excelDone = createOpenXmlDoneResult("excel");

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
  const validationXml = `<dataValidation type="${escapeXml(type)}" allowBlank="1" sqref="${escapeXml(address)}"><formula1>"${escapeXml(values)}"</formula1></dataValidation>`;
  zip.file(partName, mergeDataValidationsXml(xml, validationXml));

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

function mergeDataValidationsXml(xml: string, validationXml: string): string {
  const containerMatch = /<dataValidations\b[^>]*>([\s\S]*?)<\/dataValidations>/.exec(xml);
  if (containerMatch) {
    const openingTag = /^<dataValidations\b[^>]*>/.exec(containerMatch[0])?.[0];
    if (!openingTag) return xml;
    const count = (containerMatch[1].match(/<dataValidation\b/g) || []).length + 1;
    const updatedOpeningTag = setXmlAttribute(openingTag, "count", String(count));
    const replacement = `${updatedOpeningTag}${containerMatch[1]}${validationXml}</dataValidations>`;
    return xml.replace(containerMatch[0], replacement);
  }

  const selfClosingMatch = /<dataValidations\b[^>]*\/>/.exec(xml);
  if (selfClosingMatch) {
    const openingTag = setXmlAttribute(
      selfClosingMatch[0].replace(/\/>$/, ">"),
      "count",
      "1"
    );
    return xml.replace(
      selfClosingMatch[0],
      `${openingTag}${validationXml}</dataValidations>`
    );
  }

  return insertBeforeWorksheetEnd(
    xml,
    `<dataValidations count="1">${validationXml}</dataValidations>`
  );
}

function setXmlAttribute(tagXml: string, name: string, value: string): string {
  const attributeRe = new RegExp(`\\b${name}\\s*=\\s*["'][^"']*["']`);
  if (attributeRe.test(tagXml)) {
    return tagXml.replace(attributeRe, `${name}="${value}"`);
  }
  return tagXml.replace(/>$/, ` ${name}="${value}">`);
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

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
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
