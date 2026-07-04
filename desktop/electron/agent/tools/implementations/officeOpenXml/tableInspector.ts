/**
 * Open XML 表格检查器
 *
 * 关联模块：
 * - types.ts: 定义表格检查输入输出。
 * - officeOpenXmlFileBridge.ts: 将本模块暴露给 OfficeFileBridge。
 */

import { readFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import { decodeXmlText as unescapeXmlText } from "../../../shared/xmlEntities";
import type {
  OfficeOpenXmlDocumentType,
  OfficeOpenXmlTableCell,
  OfficeOpenXmlTableInspectInput,
  OfficeOpenXmlTableInspectResult,
  OfficeOpenXmlTableRow,
  OfficeOpenXmlTableSummary,
} from "./types";

const WORD_TABLE_PART_RE = /^word\/document\.xml$/;
const PRESENTATION_TABLE_PART_RE = /^ppt\/slides\/slide\d+\.xml$/;
const SPREADSHEET_TABLE_PART_RE = /^xl\/worksheets\/sheet\d+\.xml$/;

function detectDocumentType(filePath: string): OfficeOpenXmlDocumentType {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".docx") return "word";
  if (ext === ".pptx") return "presentation";
  if (ext === ".xlsx") return "spreadsheet";
  throw new Error(`仅支持 .docx、.pptx 和 .xlsx 文件: ${filePath}`);
}

function isTablePart(documentType: OfficeOpenXmlDocumentType, partName: string): boolean {
  if (documentType === "word") return WORD_TABLE_PART_RE.test(partName);
  if (documentType === "presentation") return PRESENTATION_TABLE_PART_RE.test(partName);
  return SPREADSHEET_TABLE_PART_RE.test(partName);
}

function blocks(xml: string, tagName: string): string[] {
  const re = new RegExp(`<${tagName}\\b[\\s\\S]*?<\\/${tagName}>`, "g");
  return Array.from(xml.matchAll(re), (match) => match[0]);
}

function firstAttribute(xml: string, name: string): string | undefined {
  const re = new RegExp(`\\b${name}="([^"]+)"`);
  return re.exec(xml)?.[1];
}

function extractText(xml: string, textTagName: string): string {
  const textRe = new RegExp(`<${textTagName}\\b[^>]*>([\\s\\S]*?)<\\/${textTagName}>`, "g");
  return Array.from(xml.matchAll(textRe), (match) => unescapeXmlText(match[1])).join("");
}

function headerGuess(cells: OfficeOpenXmlTableCell[]): boolean {
  if (cells.length === 0) return false;
  const textCells = cells.filter((cell) => cell.text && Number.isNaN(Number(cell.text)));
  return textCells.length >= Math.ceil(cells.length / 2);
}

function detectFillColor(xml: string): string | undefined {
  return firstAttribute(xml, "w:fill") || firstAttribute(xml, "fill") || firstAttribute(xml, "val");
}

function parseWordTables(xml: string, partName: string, startIndex: number): OfficeOpenXmlTableSummary[] {
  return blocks(xml, "w:tbl").map((tableXml, tableOffset) => {
    const rows = blocks(tableXml, "w:tr").map((rowXml, rowIndex) => {
      const cells = blocks(rowXml, "w:tc").map((cellXml, columnIndex) => ({
        text: extractText(cellXml, "w:t"),
        rowIndex,
        columnIndex,
        bold: /<w:b\b/.test(cellXml),
        fillColor: detectFillColor(cellXml),
        alignment: firstAttribute(cellXml, "w:val"),
      }));
      return { rowIndex, isHeaderGuess: rowIndex === 0 && headerGuess(cells), cells };
    });
    return toTableSummary(startIndex + tableOffset, partName, rows);
  });
}

function parsePresentationTables(xml: string, partName: string, startIndex: number): OfficeOpenXmlTableSummary[] {
  return blocks(xml, "a:tbl").map((tableXml, tableOffset) => {
    const rows = blocks(tableXml, "a:tr").map((rowXml, rowIndex) => {
      const cells = blocks(rowXml, "a:tc").map((cellXml, columnIndex) => ({
        text: extractText(cellXml, "a:t"),
        rowIndex,
        columnIndex,
        bold: /<a:rPr\b[^>]*\bb="1"/.test(cellXml),
        fillColor: detectFillColor(cellXml),
        alignment: firstAttribute(cellXml, "algn"),
      }));
      return { rowIndex, isHeaderGuess: rowIndex === 0 && headerGuess(cells), cells };
    });
    return toTableSummary(startIndex + tableOffset, partName, rows);
  });
}

function parseSpreadsheetTables(xml: string, partName: string, startIndex: number): OfficeOpenXmlTableSummary[] {
  const rowBlocks = blocks(xml, "row");
  if (rowBlocks.length === 0) return [];

  const rows = rowBlocks.map((rowXml, rowIndex) => {
    const cells = blocks(rowXml, "c").map((cellXml, columnIndex) => {
      const text = extractText(cellXml, "t") || extractText(cellXml, "v");
      return {
        text,
        rowIndex,
        columnIndex,
        reference: firstAttribute(cellXml, "r"),
        fillColor: detectFillColor(cellXml),
        alignment: firstAttribute(cellXml, "horizontal"),
      };
    });
    return { rowIndex, isHeaderGuess: rowIndex === 0 && headerGuess(cells), cells };
  });

  return [toTableSummary(startIndex, partName, rows)];
}

function toTableSummary(index: number, partName: string, rows: OfficeOpenXmlTableRow[]): OfficeOpenXmlTableSummary {
  return {
    index,
    partName,
    rows,
    columns: Math.max(0, ...rows.map((row) => row.cells.length)),
  };
}

export async function inspectOfficeOpenXmlTables(
  input: OfficeOpenXmlTableInspectInput
): Promise<OfficeOpenXmlTableInspectResult> {
  const documentType = detectDocumentType(input.filePath);
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const tables: OfficeOpenXmlTableSummary[] = [];

  for (const partName of Object.keys(zip.files).filter((name) => isTablePart(documentType, name)).sort()) {
    const file = zip.file(partName);
    if (!file) continue;
    const xml = await file.async("text");
    const startIndex = tables.length;
    if (documentType === "word") tables.push(...parseWordTables(xml, partName, startIndex));
    if (documentType === "presentation") tables.push(...parsePresentationTables(xml, partName, startIndex));
    if (documentType === "spreadsheet") tables.push(...parseSpreadsheetTables(xml, partName, startIndex));
  }

  return {
    engine: "openxml",
    operation: "inspectTable",
    documentType,
    filePath: input.filePath,
    target: input.target,
    tableCount: tables.length,
    tables,
  };
}
