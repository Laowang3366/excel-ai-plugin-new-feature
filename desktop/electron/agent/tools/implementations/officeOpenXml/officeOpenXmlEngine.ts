/**
 * Open XML Office 文件编辑引擎
 *
 * 直接读写 docx/pptx/xlsx 的 ZIP + XML 结构，不依赖 Office COM、PowerShell 或已打开的 Office 应用。
 *
 * 关联模块：
 * - types.ts: 引擎输入输出契约。
 * - tools/executors/officeExecutors.ts: 将本引擎暴露成模型可调用工具。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import { decodeXmlText as unescapeXmlText, escapeXmlText } from "../../../shared/xmlEntities";
import type {
  OfficeOpenXmlDocumentType,
  OfficeOpenXmlInspectResult,
  OfficeOpenXmlReplaceInput,
  OfficeOpenXmlReplaceResult,
  OfficeOpenXmlTextPart,
} from "./types";

const WORD_TEXT_PART_RE = /^word\/(?:document|header\d+|footer\d+)\.xml$/;
const PRESENTATION_TEXT_PART_RE = /^ppt\/slides\/slide\d+\.xml$/;
const SPREADSHEET_TEXT_PART_RE = /^(?:xl\/sharedStrings\.xml|xl\/worksheets\/sheet\d+\.xml)$/;

function detectDocumentType(filePath: string): OfficeOpenXmlDocumentType {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".docx") return "word";
  if (ext === ".pptx") return "presentation";
  if (ext === ".xlsx") return "spreadsheet";
  throw new Error(`仅支持 .docx、.pptx 和 .xlsx 文件: ${filePath}`);
}

function isTextPart(documentType: OfficeOpenXmlDocumentType, partName: string): boolean {
  if (documentType === "word") return WORD_TEXT_PART_RE.test(partName);
  if (documentType === "presentation") return PRESENTATION_TEXT_PART_RE.test(partName);
  return SPREADSHEET_TEXT_PART_RE.test(partName);
}

function textTagName(documentType: OfficeOpenXmlDocumentType): string {
  if (documentType === "word") return "w:t";
  if (documentType === "presentation") return "a:t";
  return "t";
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractTextFromXml(documentType: OfficeOpenXmlDocumentType, xml: string): string {
  const tagName = textTagName(documentType);
  const re = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)<\\/${tagName}>`, "g");
  const values: string[] = [];
  for (const match of xml.matchAll(re)) {
    values.push(unescapeXmlText(match[1]));
  }
  return values.join("\n");
}

function replaceEscapedTextInXml(
  documentType: OfficeOpenXmlDocumentType,
  xml: string,
  findText: string,
  replaceText: string,
  matchCase: boolean
): { xml: string; replacements: number } {
  const escapedFind = escapeXmlText(findText);
  if (!escapedFind) {
    throw new Error("findText 不能为空");
  }
  const tagName = textTagName(documentType);
  const textRe = new RegExp(`(<${tagName}[^>]*>)([\\s\\S]*?)(<\\/${tagName}>)`, "g");
  const findRe = new RegExp(escapeRegExp(escapedFind), matchCase ? "g" : "gi");
  let replacements = 0;
  const nextXml = xml.replace(textRe, (_match, openTag: string, innerText: string, closeTag: string) => {
    const nextInnerText = innerText.replace(findRe, () => {
      replacements += 1;
      return escapeXmlText(replaceText);
    });
    return openTag + nextInnerText + closeTag;
  });
  return { xml: nextXml, replacements };
}

async function loadPackage(filePath: string): Promise<{ zip: JSZip; documentType: OfficeOpenXmlDocumentType }> {
  const documentType = detectDocumentType(filePath);
  const zip = await JSZip.loadAsync(await readFile(filePath));
  return { zip, documentType };
}

async function readTextParts(zip: JSZip, documentType: OfficeOpenXmlDocumentType): Promise<OfficeOpenXmlTextPart[]> {
  const parts: OfficeOpenXmlTextPart[] = [];
  const names = Object.keys(zip.files).filter((partName) => isTextPart(documentType, partName)).sort();

  for (const partName of names) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    const text = extractTextFromXml(documentType, xml);
    if (!text) continue;
    parts.push({ partName, text, textLength: text.length });
  }

  return parts;
}

function defaultOutputPath(filePath: string): string {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);
  return path.join(dir, `${base}-edited${ext}`);
}

export async function inspectOfficeOpenXmlFile(filePath: string): Promise<OfficeOpenXmlInspectResult> {
  const { zip, documentType } = await loadPackage(filePath);
  const textParts = await readTextParts(zip, documentType);
  const fullText = textParts.map((part) => part.text).join("\n");

  return {
    engine: "openxml",
    operation: "inspect",
    documentType,
    filePath,
    textPartCount: textParts.length,
    textCharCount: fullText.length,
    textPreview: fullText.slice(0, 4000),
    textParts,
  };
}

export async function replaceOfficeOpenXmlText(input: OfficeOpenXmlReplaceInput): Promise<OfficeOpenXmlReplaceResult> {
  const { zip, documentType } = await loadPackage(input.filePath);
  const changedParts: OfficeOpenXmlReplaceResult["changedParts"] = [];
  const partNames = Object.keys(zip.files).filter((partName) => isTextPart(documentType, partName)).sort();
  let replacements = 0;

  for (const partName of partNames) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    const replaced = replaceEscapedTextInXml(documentType, xml, input.findText, input.replaceText, input.matchCase ?? false);
    if (replaced.replacements > 0) {
      zip.file(partName, replaced.xml);
      changedParts.push({ partName, replacements: replaced.replacements });
      replacements += replaced.replacements;
    }
  }

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));

  return {
    engine: "openxml",
    operation: "replaceText",
    documentType,
    filePath: input.filePath,
    outputPath,
    findText: input.findText,
    replaceText: input.replaceText,
    replacements,
    changedParts,
  };
}
