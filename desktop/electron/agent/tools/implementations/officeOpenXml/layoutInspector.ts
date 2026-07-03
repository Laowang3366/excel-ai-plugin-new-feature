/**
 * Open XML 布局检查器
 *
 * 关联模块：
 * - types.ts: 定义布局检查输入输出。
 * - officeOpenXmlFileBridge.ts: 将本模块暴露给 OfficeFileBridge。
 */

import { readFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import type {
  OfficeOpenXmlDocumentType,
  OfficeOpenXmlLayoutInspectInput,
  OfficeOpenXmlLayoutInspectResult,
  OfficeOpenXmlLayoutObject,
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

function unescapeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function extractTextValues(documentType: OfficeOpenXmlDocumentType, xml: string): string[] {
  const tagName = textTagName(documentType);
  const re = new RegExp(`<${tagName}[^>]*>([\\s\\S]*?)<\\/${tagName}>`, "g");
  return Array.from(xml.matchAll(re), (match) => unescapeXmlText(match[1])).filter(Boolean);
}

export async function inspectOfficeOpenXmlLayout(
  input: OfficeOpenXmlLayoutInspectInput
): Promise<OfficeOpenXmlLayoutInspectResult> {
  const documentType = detectDocumentType(input.filePath);
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const objects: OfficeOpenXmlLayoutObject[] = [];

  for (const partName of Object.keys(zip.files).filter((name) => isTextPart(documentType, name)).sort()) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    for (const text of extractTextValues(documentType, xml)) {
      objects.push({ type: "text", partName, text, textLength: text.length });
    }
  }

  return {
    engine: "openxml",
    operation: "inspectLayout",
    documentType,
    filePath: input.filePath,
    target: input.target,
    objectCount: objects.length,
    objects,
  };
}
