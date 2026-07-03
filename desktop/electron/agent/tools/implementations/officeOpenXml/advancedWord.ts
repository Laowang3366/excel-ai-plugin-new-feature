/**
 * Word Open XML 高级操作
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: 将统一 Office action 路由到本模块。
 * - tableStyler.ts: 复用 Word 表格样式处理。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import { doneResult, failedResult, needsComResult, unsupportedResult } from "../../officeCore/results";
import type { OfficeActionKind, OfficeActionResult } from "../../officeCore/types";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";
import type { OfficeOpenXmlTableStylePreset } from "./types";

export interface WordAdvancedActionInput {
  operation: string;
  filePath: string;
  outputPath?: string;
  target?: string;
  action?: OfficeActionKind;
  params?: Record<string, unknown>;
}

const WORD_DOCUMENT_PART = "word/document.xml";
const TABLE_STYLES = new Set<OfficeOpenXmlTableStylePreset>(["professional", "compact", "financial"]);

export async function applyWordAdvancedAction(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  try {
    if (input.operation === "createDocument") {
      return await createDocument(input);
    }
    if (input.operation === "applyHeadingStyles") {
      return await applyHeadingStyles(input);
    }
    if (input.operation === "styleTables") {
      return await styleTables(input);
    }
    if (input.operation === "setHeaderFooter") {
      return await setHeaderFooter(input);
    }
    if (input.operation === "insertOrUpdateToc" || input.operation === "insertOrReplaceImage") {
      return needsComResult({
        app: "word",
        action: input.action || "insert",
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: `${input.operation} 需要 Word 刷新字段或维护关系图，需显式 COM 兜底`,
      });
    }

    return unsupportedResult({
      app: "word",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: `暂不支持 Word Open XML 高级操作: ${input.operation}`,
    });
  } catch (error) {
    return failedResult({
      app: "word",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
    }, error);
  }
}

async function createDocument(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const outputPath = input.outputPath || input.filePath;
  const title = typeof input.params?.title === "string" ? input.params.title : "";
  const body = normalizeParagraphs(input.params?.paragraphs ?? input.params?.text ?? input.params?.body);
  const zip = new JSZip();

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("word/styles.xml", basicStylesXml());
  zip.file(WORD_DOCUMENT_PART, documentXml(title, body));

  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return wordDone(input, outputPath, [WORD_DOCUMENT_PART, "word/styles.xml"], "已使用内置 Open XML 创建 Word 文档", {
    engine: "openxml",
    operation: "createDocument",
    filePath: input.filePath,
    outputPath,
    title,
    paragraphCount: body.length,
  });
}

async function applyHeadingStyles(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const part = zip.file(WORD_DOCUMENT_PART);
  if (!part) throw new Error(`找不到 Word 正文部件: ${WORD_DOCUMENT_PART}`);

  const startsWith = typeof input.params?.startsWith === "string" ? input.params.startsWith : "";
  const level = normalizeHeadingLevel(input.params?.level);
  const xml = await part.async("text");
  let changed = false;
  const nextXml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    const text = paragraphText(paragraphXml);
    if (startsWith && !text.startsWith(startsWith)) return paragraphXml;
    changed = true;
    return applyParagraphStyle(paragraphXml, `Heading${level}`);
  });
  zip.file(WORD_DOCUMENT_PART, nextXml);

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return wordDone(input, outputPath, changed ? [WORD_DOCUMENT_PART] : [], "已应用 Word 标题样式");
}

async function styleTables(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const style = normalizeTableStyle(input.params?.style);
  const result = await applyOfficeOpenXmlTableStyle({
    filePath: input.filePath,
    outputPath: input.outputPath,
    target: input.target,
    style,
  });
  return wordDone(input, result.outputPath, result.changedParts, "已应用 Word 表格样式", result);
}

async function setHeaderFooter(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const kind = input.operation === "setHeaderFooter" && input.params?.kind === "footer" ? "footer" : "header";
  const partName = kind === "footer" ? "word/footer1.xml" : "word/header1.xml";
  const text = typeof input.params?.text === "string" ? input.params.text : "";
  zip.file(partName, wordTextPartXml(kind, text));

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return wordDone(input, outputPath, [partName], kind === "footer" ? "已写入 Word 页脚" : "已写入 Word 页眉");
}

function wordDone(
  input: WordAdvancedActionInput,
  outputPath: string,
  changedParts: string[],
  summary: string,
  data?: unknown
): OfficeActionResult {
  return doneResult({
    engine: "openxml",
    app: "word",
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

function paragraphText(paragraphXml: string): string {
  return [...paragraphXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function applyParagraphStyle(paragraphXml: string, styleId: string): string {
  const styleXml = `<w:pStyle w:val="${styleId}" />`;
  if (/<w:pPr\b[\s\S]*?<\/w:pPr>/.test(paragraphXml)) {
    return paragraphXml.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/, (pPrXml) => {
      if (/<w:pStyle\b[^>]*(?:\/>|><\/w:pStyle>)/.test(pPrXml)) {
        return pPrXml.replace(/<w:pStyle\b[^>]*(?:\/>|><\/w:pStyle>)/, styleXml);
      }
      return pPrXml.replace(/<w:pPr\b[^>]*>/, (tag) => `${tag}${styleXml}`);
    });
  }
  return paragraphXml.replace(/<w:p\b[^>]*>/, (tag) => `${tag}<w:pPr>${styleXml}</w:pPr>`);
}

function wordTextPartXml(kind: "header" | "footer", text: string): string {
  const root = kind === "footer" ? "w:ftr" : "w:hdr";
  return `<${root}><w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p></${root}>`;
}

function normalizeParagraphs(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.map((item) => String(item)).filter((item) => item.length > 0);
  }
  if (typeof value === "string" && value.length > 0) {
    return value.split(/\r?\n/).filter((item) => item.length > 0);
  }
  return [];
}

function documentXml(title: string, paragraphs: string[]): string {
  const parts = [
    title ? paragraphXml(title, "Title") : "",
    ...paragraphs.map((text) => paragraphXml(text)),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${parts}<w:sectPr /></w:body>
</w:document>`;
}

function paragraphXml(text: string, styleId?: string): string {
  const style = styleId ? `<w:pPr><w:pStyle w:val="${styleId}" /></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function basicStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>`;
}

function normalizeHeadingLevel(value: unknown): number {
  return typeof value === "number" && value >= 1 && value <= 9 ? Math.floor(value) : 1;
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

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function unescapeXml(value: string): string {
  return value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}
