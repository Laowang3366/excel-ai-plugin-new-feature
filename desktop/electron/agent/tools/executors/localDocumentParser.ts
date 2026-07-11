/**
 * Local document parser for ocr.parseDocument.
 *
 * Provides a free first-pass parser for text, CSV and Open XML Office files.
 * Files that need real OCR, such as scanned PDFs and images, are reported as
 * unsupported so the caller can decide whether to use a remote OCR service.
 */

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { DocumentParser } from "../../knowledge/documentParser";
import { extractOpenXmlTextValues, readOpenXmlTextParts } from "../../shared/openXmlText";
import { extractMarkdownTables } from "../../../shared/markdownTables";

export interface LocalParsedDocument {
  filename: string;
  text: string;
  rows: string[][];
  provider: "local";
  sourceType: string;
  warnings: string[];
  error?: string;
}

const TEXT_EXTENSIONS = new Set([".txt", ".md", ".csv"]);
const KNOWLEDGE_PARSER_EXTENSIONS = new Set([".xlsx", ".xlsm", ".csv", ".md", ".txt"]);
const LOCAL_UNSUPPORTED_EXTENSIONS = new Set([
  ".png",
  ".jpg",
  ".jpeg",
  ".webp",
  ".bmp",
  ".tif",
  ".tiff",
  ".pdf",
  ".doc",
  ".ppt",
  ".xls",
]);

export function isLocallyUnsupportedForOcr(filePath: string): boolean {
  return LOCAL_UNSUPPORTED_EXTENSIONS.has(path.extname(filePath).toLowerCase());
}

export async function parseFilesLocally(filePaths: string[]): Promise<LocalParsedDocument[]> {
  return Promise.all(filePaths.map((filePath) => parseFileLocally(filePath)));
}

async function parseFileLocally(filePath: string): Promise<LocalParsedDocument> {
  const filename = path.basename(filePath);
  const ext = path.extname(filePath).toLowerCase();

  if (KNOWLEDGE_PARSER_EXTENSIONS.has(ext)) {
    return parseWithKnowledgeParser(filePath, filename, ext);
  }
  if (ext === ".docx") {
    return parseDocx(filePath, filename);
  }
  if (ext === ".pptx") {
    return parsePptx(filePath, filename);
  }
  if (isLocallyUnsupportedForOcr(filePath)) {
    return {
      filename,
      text: "",
      rows: [],
      provider: "local",
      sourceType: ext.replace(/^\./, "") || "unknown",
      warnings: [
        `${filename}: 本地免费解析无法直接识别该类型的可见图像内容，需要远程 OCR、视觉模型或对应 Office/文件工具兜底。`,
      ],
      error: "local_ocr_unsupported",
    };
  }

  return {
    filename,
    text: "",
    rows: [],
    provider: "local",
    sourceType: ext.replace(/^\./, "") || "unknown",
    warnings: [`${filename}: 当前本地解析器暂不支持该文件类型。`],
    error: "local_unsupported",
  };
}

async function parseWithKnowledgeParser(
  filePath: string,
  filename: string,
  ext: string
): Promise<LocalParsedDocument> {
  try {
    const parser = new DocumentParser();
    const chunks = await parser.parseAsync(filePath);
    const text = chunks.map((chunk) => {
      const title = chunk.metadata.sheetName
        ? `### ${chunk.metadata.sheetName}`
        : `### ${chunk.sourceName}`;
      return `${title}\n${chunk.content}`;
    }).join("\n\n");
    const rows = chunks.flatMap((chunk) => extractRowsFromChunk(chunk.content));
    const warnings: string[] = [];
    if (!text.trim()) {
      warnings.push(`${filename}: 本地解析未提取到可用文本。`);
    }
    return {
      filename,
      text,
      rows,
      provider: "local",
      sourceType: ext.replace(/^\./, ""),
      warnings,
      error: text.trim() ? undefined : "local_empty",
    };
  } catch (error: any) {
    if (TEXT_EXTENSIONS.has(ext)) {
      return parsePlainTextFallback(filePath, filename, ext, error);
    }
    return {
      filename,
      text: "",
      rows: [],
      provider: "local",
      sourceType: ext.replace(/^\./, ""),
      warnings: [`${filename}: 本地解析失败: ${error?.message || String(error)}`],
      error: error?.message || "local_parse_failed",
    };
  }
}

function parsePlainTextFallback(
  filePath: string,
  filename: string,
  ext: string,
  originalError: unknown
): LocalParsedDocument {
  try {
    const text = fs.readFileSync(filePath, "utf8");
    const rows = ext === ".csv" ? parseCsvRows(text) : extractMarkdownTables(text);
    return {
      filename,
      text,
      rows,
      provider: "local",
      sourceType: ext.replace(/^\./, ""),
      warnings: [],
      error: text.trim() ? undefined : "local_empty",
    };
  } catch (fallbackError: any) {
    return {
      filename,
      text: "",
      rows: [],
      provider: "local",
      sourceType: ext.replace(/^\./, ""),
      warnings: [
        `${filename}: 本地文本读取失败: ${fallbackError?.message || String(fallbackError)}；原解析错误: ${
          originalError instanceof Error ? originalError.message : String(originalError)
        }`,
      ],
      error: fallbackError?.message || "local_text_parse_failed",
    };
  }
}

async function parseDocx(filePath: string, filename: string): Promise<LocalParsedDocument> {
  try {
    const zip = await JSZip.loadAsync(await fs.promises.readFile(filePath));
    const documentText = (await readOpenXmlTextParts(zip, /^word\/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml$/i, {
      tagName: "w:t",
      includeEmpty: false,
    })).map((part) => part.text).join("\n");
    const rows = await readDocxTables(zip);
    const textParts = [documentText.trim()];
    if (rows.length > 0) {
      textParts.push(formatRowsAsMarkdown(rows));
    }
    const text = textParts.filter(Boolean).join("\n\n");
    return {
      filename,
      text,
      rows,
      provider: "local",
      sourceType: "docx",
      warnings: text.trim() ? [] : [`${filename}: 本地 DOCX 解析未提取到可用文本。`],
      error: text.trim() ? undefined : "local_empty",
    };
  } catch (error: any) {
    return {
      filename,
      text: "",
      rows: [],
      provider: "local",
      sourceType: "docx",
      warnings: [`${filename}: 本地 DOCX 解析失败: ${error?.message || String(error)}`],
      error: error?.message || "local_docx_parse_failed",
    };
  }
}

async function parsePptx(filePath: string, filename: string): Promise<LocalParsedDocument> {
  try {
    const zip = await JSZip.loadAsync(await fs.promises.readFile(filePath));
    const slideNames = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((a, b) => slideNumber(a) - slideNumber(b));
    const slideTexts: string[] = [];

    for (const slideName of slideNames) {
      const slide = zip.file(slideName);
      if (!slide) continue;
      const xml = await slide.async("text");
      const values = extractOpenXmlTextValues(xml, { tagName: "a:t", includeEmpty: false });
      if (values.length > 0) {
        slideTexts.push(`### Slide ${slideNumber(slideName)}\n${values.join("\n")}`);
      }
    }

    const text = slideTexts.join("\n\n");
    return {
      filename,
      text,
      rows: [],
      provider: "local",
      sourceType: "pptx",
      warnings: text.trim() ? [] : [`${filename}: 本地 PPTX 解析未提取到可用文本。`],
      error: text.trim() ? undefined : "local_empty",
    };
  } catch (error: any) {
    return {
      filename,
      text: "",
      rows: [],
      provider: "local",
      sourceType: "pptx",
      warnings: [`${filename}: 本地 PPTX 解析失败: ${error?.message || String(error)}`],
      error: error?.message || "local_pptx_parse_failed",
    };
  }
}

async function readDocxTables(zip: JSZip): Promise<string[][]> {
  const part = zip.file("word/document.xml");
  if (!part) return [];
  const xml = await part.async("text");
  const rows: string[][] = [];
  for (const tableMatch of xml.matchAll(/<w:tbl\b[\s\S]*?<\/w:tbl>/g)) {
    for (const rowMatch of tableMatch[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) {
      const cells = Array.from(rowMatch[0].matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g), (cellMatch) =>
        extractOpenXmlTextValues(cellMatch[0], { tagName: "w:t" }).join("")
      );
      if (cells.some((cell) => cell.trim())) {
        rows.push(cells);
      }
    }
  }
  return rows;
}

function extractRowsFromChunk(content: string): string[][] {
  const rows = extractMarkdownTables(content);
  if (rows.length > 0) return rows;

  const lines = content.split(/\r?\n/).filter((line) => line.trim());
  const headerIndex = lines.findIndex((line) => line.startsWith("【表头】"));
  if (headerIndex < 0) return [];

  const parsedRows = lines.slice(headerIndex).map((line, index) => {
    const normalized = index === 0 ? line.replace(/^【表头】\s*/, "") : line;
    return normalized.split("|").map((cell) => cell.trim());
  });
  return parsedRows.filter((row) => row.some(Boolean));
}

function parseCsvRows(text: string): string[][] {
  return text
    .split(/\r?\n/)
    .filter((line) => line.trim())
    .map((line) => line.split(",").map((cell) => cell.trim().replace(/^"(.*)"$/, "$1")));
}

function formatRowsAsMarkdown(rows: string[][]): string {
  if (rows.length === 0) return "";
  const width = Math.max(...rows.map((row) => row.length));
  const normalized = rows.map((row) => Array.from({ length: width }, (_, index) => row[index] || ""));
  const header = normalized[0];
  const separator = Array.from({ length: width }, () => "---");
  const body = normalized.slice(1);
  return [
    "| " + header.join(" | ") + " |",
    "| " + separator.join(" | ") + " |",
    ...body.map((row) => "| " + row.join(" | ") + " |"),
  ].join("\n");
}

function slideNumber(partName: string): number {
  return Number(/slide(\d+)\.xml$/i.exec(partName)?.[1] || 0);
}
