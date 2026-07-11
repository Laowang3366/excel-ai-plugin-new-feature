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
import { extractOpenXmlText } from "../../../shared/openXmlText";
import { escapeXmlText } from "../../../shared/xmlEntities";
import {
  detectOfficeOpenXmlDocumentType,
  getOfficeOpenXmlTextTagName,
  isOfficeOpenXmlTextPart,
} from "./documentParts";
import type {
  OfficeOpenXmlDocumentType,
  OfficeOpenXmlInspectResult,
  OfficeOpenXmlReplaceInput,
  OfficeOpenXmlReplaceResult,
  OfficeOpenXmlTextPart,
} from "./types";

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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
  const tagName = getOfficeOpenXmlTextTagName(documentType);
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
  const documentType = detectOfficeOpenXmlDocumentType(filePath);
  const zip = await JSZip.loadAsync(await readFile(filePath));
  return { zip, documentType };
}

async function readTextParts(zip: JSZip, documentType: OfficeOpenXmlDocumentType): Promise<OfficeOpenXmlTextPart[]> {
  const parts: OfficeOpenXmlTextPart[] = [];
  const names = Object.keys(zip.files).filter((partName) => isOfficeOpenXmlTextPart(documentType, partName)).sort();

  for (const partName of names) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    const text = extractOpenXmlText(xml, { tagName: getOfficeOpenXmlTextTagName(documentType) });
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
  const partNames = Object.keys(zip.files).filter((partName) => isOfficeOpenXmlTextPart(documentType, partName)).sort();
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
