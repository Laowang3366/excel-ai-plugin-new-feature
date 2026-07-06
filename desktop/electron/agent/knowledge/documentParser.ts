/**
 * 文档解析器
 *
 * 支持解析 .xlsx / .xlsm / .csv / .json / .docx / .pptx / .md / .txt 文件。
 * Office 文件使用内置 Open XML ZIP/XML 读取，避免引入原生二进制依赖。
 */

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { extractOpenXmlParagraphTexts, extractOpenXmlTextValues } from "../shared/openXmlText";
import { decodeXmlText } from "../shared/xmlEntities";
import { flattenJson } from "./jsonFlatten";
import type { KnowledgeFileType } from "./types";

const MAX_EXCEL_PARSE_BYTES = 25 * 1024 * 1024;
const MAX_EXCEL_DATA_ROWS = 500;
const MAX_JSON_LINES = 2000;

interface WorkbookSheetInfo {
  name: string;
  partName: string;
}

interface ParsedWorksheetRows {
  rows: string[][];
  totalRows: number;
  maxCol: number;
  maxRow: number;
}

export interface RawChunk {
  content: string;
  sourcePath: string;
  sourceName: string;
  sourceType: KnowledgeFileType;
  metadata: Record<string, unknown> & {
    sheetName?: string;
    tableRange?: string;
    headers?: string[];
    rowCount?: number;
    colCount?: number;
  };
}

export class DocumentParser {
  parse(filePath: string): RawChunk[] {
    const ext = this.getFileType(filePath);
    const sourceName = path.basename(filePath);

    switch (ext) {
      case "csv":
        return this.parseCsv(filePath, sourceName);
      case "json":
        return this.parseJson(filePath, sourceName);
      case "md":
        return this.parseMarkdown(filePath, sourceName);
      case "txt":
        return this.parseText(filePath, sourceName);
      case "xlsx":
      case "xlsm":
        throw new Error("Excel 文件解析需要使用 parseAsync");
      case "docx":
      case "pptx":
        throw new Error("Office Open XML 文件解析需要使用 parseAsync");
      default:
        throw new Error(`不支持的文件类型: ${ext}`);
    }
  }

  async parseAsync(filePath: string): Promise<RawChunk[]> {
    const ext = this.getFileType(filePath);
    const sourceName = path.basename(filePath);

    switch (ext) {
      case "xlsx":
      case "xlsm":
        return await this.parseExcel(filePath, sourceName, ext);
      case "docx":
        return await this.parseDocx(filePath, sourceName);
      case "pptx":
        return await this.parsePptx(filePath, sourceName);
      case "csv":
      case "json":
      case "md":
      case "txt":
        return this.parse(filePath);
      default:
        throw new Error(`不支持的文件类型: ${ext}`);
    }
  }

  isSupported(filePath: string): boolean {
    try {
      const ext = this.getFileType(filePath);
      return ["xlsx", "xlsm", "csv", "json", "docx", "pptx", "md", "txt"].includes(ext);
    } catch {
      return false;
    }
  }

  getSupportedExtensions(): string[] {
    return ["xlsx", "xlsm", "csv", "json", "docx", "pptx", "md", "txt"];
  }

  private async parseExcel(
    filePath: string,
    sourceName: string,
    sourceType: KnowledgeFileType
  ): Promise<RawChunk[]> {
    const stats = fs.statSync(filePath);
    if (stats.size > MAX_EXCEL_PARSE_BYTES) {
      throw new Error(`Excel 文件过大，知识库索引最多支持 ${Math.floor(MAX_EXCEL_PARSE_BYTES / 1024 / 1024)}MB: ${sourceName}`);
    }

    const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
    const sharedStrings = await this.readSharedStrings(zip);
    const sheets = await this.readWorkbookSheets(zip);
    const chunks: RawChunk[] = [];

    for (const sheetInfo of sheets) {
      const sheetPart = zip.file(sheetInfo.partName);
      if (!sheetPart) continue;

      const xml = await sheetPart.async("text");
      const parsedRows = this.parseWorksheetRows(xml, sharedStrings, MAX_EXCEL_DATA_ROWS + 1);
      if (parsedRows.rows.length === 0) continue;

      const headers = parsedRows.rows[0].map((h) => h.trim());
      const rowCount = Math.max(0, parsedRows.totalRows - 1);
      const colCount = Math.max(headers.length, parsedRows.maxCol);
      const lines: string[] = [`【表头】${headers.join(" | ")}`];

      for (const row of parsedRows.rows.slice(1, MAX_EXCEL_DATA_ROWS + 1)) {
        lines.push(row.map((cell) => cell.trim()).join(" | "));
      }

      if (rowCount > MAX_EXCEL_DATA_ROWS) {
        lines.push(`...（还有 ${rowCount - MAX_EXCEL_DATA_ROWS} 行未展示）`);
      }

      chunks.push({
        content: lines.join("\n"),
        sourcePath: filePath,
        sourceName,
        sourceType,
        metadata: {
          sheetName: sheetInfo.name,
          tableRange: this.detectWorksheetRange(xml, parsedRows),
          headers,
          rowCount,
          colCount,
        },
      });
    }

    return chunks;
  }

  private parseCsv(filePath: string, sourceName: string): RawChunk[] {
    const content = fs.readFileSync(filePath, "utf-8");
    const lines = content.split("\n").filter((line) => line.trim().length > 0);
    if (lines.length === 0) return [];

    const parseLine = (line: string): string[] =>
      line.split(",").map((cell) => cell.trim().replace(/^"(.*)"$/, "$1"));

    const headers = parseLine(lines[0]);
    const rowCount = lines.length - 1;
    const textLines: string[] = [`【表头】${headers.join(" | ")}`];

    for (let i = 1; i < Math.min(lines.length, MAX_EXCEL_DATA_ROWS + 1); i++) {
      textLines.push(parseLine(lines[i]).join(" | "));
    }

    if (rowCount > MAX_EXCEL_DATA_ROWS) {
      textLines.push(`...（还有 ${rowCount - MAX_EXCEL_DATA_ROWS} 行未展示）`);
    }

    return [
      {
        content: textLines.join("\n"),
        sourcePath: filePath,
        sourceName,
        sourceType: "csv",
        metadata: {
          headers,
          rowCount,
          colCount: headers.length,
        },
      },
    ];
  }

  private parseJson(filePath: string, sourceName: string): RawChunk[] {
    const content = fs.readFileSync(filePath, "utf-8");
    if (!content.trim()) return [];

    let parsed: unknown;
    try {
      parsed = JSON.parse(content);
    } catch (error: any) {
      throw new Error(`JSON 解析失败: ${sourceName} (${error?.message || "格式错误"})`);
    }

    const allLines = flattenJson(parsed);
    const lines = allLines.slice(0, MAX_JSON_LINES);
    if (allLines.length > MAX_JSON_LINES) {
      lines.push(`...（还有 ${allLines.length - MAX_JSON_LINES} 个 JSON 字段未展示）`);
    }

    return [
      {
        content: lines.join("\n"),
        sourcePath: filePath,
        sourceName,
        sourceType: "json",
        metadata: {
          rowCount: allLines.length,
        },
      },
    ];
  }

  private async parseDocx(filePath: string, sourceName: string): Promise<RawChunk[]> {
    const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
    const documentPart = zip.file("word/document.xml");
    if (!documentPart) return [];

    const xml = await documentPart.async("text");
    const lines = extractOpenXmlParagraphTexts(xml);
    if (lines.length === 0) return [];

    return [
      {
        content: lines.join("\n"),
        sourcePath: filePath,
        sourceName,
        sourceType: "docx",
        metadata: {
          rowCount: lines.length,
        },
      },
    ];
  }

  private async parsePptx(filePath: string, sourceName: string): Promise<RawChunk[]> {
    const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
    const slideParts = Object.keys(zip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort((a, b) => this.slidePartNumber(a) - this.slidePartNumber(b));
    const chunks: RawChunk[] = [];

    for (const partName of slideParts) {
      const part = zip.file(partName);
      if (!part) continue;
      const xml = await part.async("text");
      const lines = extractOpenXmlTextValues(xml, { namespaceAgnostic: true })
        .map((value) => value.trim())
        .filter(Boolean);
      if (lines.length === 0) continue;
      const slideNumber = this.slidePartNumber(partName) || chunks.length + 1;
      chunks.push({
        content: [`【幻灯片 ${slideNumber}】`, ...lines].join("\n"),
        sourcePath: filePath,
        sourceName,
        sourceType: "pptx",
        metadata: {
          slideNumber,
          rowCount: lines.length,
        },
      });
    }

    return chunks;
  }

  private parseMarkdown(filePath: string, sourceName: string): RawChunk[] {
    const content = fs.readFileSync(filePath, "utf-8");
    if (!content.trim()) return [];

    return [
      {
        content,
        sourcePath: filePath,
        sourceName,
        sourceType: "md",
        metadata: {
          rowCount: content.split("\n").length,
        },
      },
    ];
  }

  private parseText(filePath: string, sourceName: string): RawChunk[] {
    const content = fs.readFileSync(filePath, "utf-8");
    if (!content.trim()) return [];

    return [
      {
        content,
        sourcePath: filePath,
        sourceName,
        sourceType: "txt",
        metadata: {
          rowCount: content.split("\n").length,
        },
      },
    ];
  }

  private async readWorkbookSheets(zip: JSZip): Promise<WorkbookSheetInfo[]> {
    const workbookPart = zip.file("xl/workbook.xml");
    const relsPart = zip.file("xl/_rels/workbook.xml.rels");
    if (!workbookPart || !relsPart) return this.fallbackWorksheetParts(zip);

    const workbookXml = await workbookPart.async("text");
    const relsXml = await relsPart.async("text");
    const relationshipTargets = this.parseWorkbookRelationships(relsXml);
    const sheets: WorkbookSheetInfo[] = [];
    const sheetRe = /<sheet\b[^>]*\/?>/g;
    let match: RegExpExecArray | null;

    while ((match = sheetRe.exec(workbookXml))) {
      const attrs = this.parseXmlAttributes(match[0]);
      const relId = attrs["r:id"] || attrs.id;
      const target = relId ? relationshipTargets.get(relId) : undefined;
      if (!target) continue;
      sheets.push({
        name: decodeXmlText(attrs.name || `Sheet${sheets.length + 1}`),
        partName: target,
      });
    }

    return sheets.length > 0 ? sheets : this.fallbackWorksheetParts(zip);
  }

  private parseWorkbookRelationships(relsXml: string): Map<string, string> {
    const targets = new Map<string, string>();
    const relRe = /<Relationship\b[^>]*\/?>/g;
    let match: RegExpExecArray | null;
    while ((match = relRe.exec(relsXml))) {
      const attrs = this.parseXmlAttributes(match[0]);
      if (attrs.Id && attrs.Target) {
        targets.set(attrs.Id, this.normalizeWorkbookTarget(attrs.Target));
      }
    }
    return targets;
  }

  private fallbackWorksheetParts(zip: JSZip): WorkbookSheetInfo[] {
    return Object.keys(zip.files)
      .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
      .sort((a, b) => this.sheetPartNumber(a) - this.sheetPartNumber(b))
      .map((partName, index) => ({ name: `Sheet${index + 1}`, partName }));
  }

  private async readSharedStrings(zip: JSZip): Promise<string[]> {
    const part = zip.file("xl/sharedStrings.xml");
    if (!part) return [];

    const xml = await part.async("text");
    const strings: string[] = [];
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let match: RegExpExecArray | null;
    while ((match = siRe.exec(xml))) {
      strings.push(extractOpenXmlTextValues(match[1], { namespaceAgnostic: true }).join(""));
    }
    return strings;
  }

  private parseWorksheetRows(
    xml: string,
    sharedStrings: string[],
    rowLimit: number
  ): ParsedWorksheetRows {
    const rows: string[][] = [];
    let totalRows = 0;
    let maxCol = 0;
    let maxRow = 0;
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
    let rowMatch: RegExpExecArray | null;

    while ((rowMatch = rowRe.exec(xml))) {
      totalRows += 1;
      const attrs = this.parseXmlAttributes(rowMatch[1]);
      const rowNumber = Number(attrs.r) || totalRows;
      maxRow = Math.max(maxRow, rowNumber);

      const parsed = this.parseWorksheetCells(rowMatch[2], sharedStrings);
      maxCol = Math.max(maxCol, parsed.length);
      if (rows.length < rowLimit) {
        rows.push(parsed);
      }
    }

    return { rows, totalRows, maxCol, maxRow };
  }

  private parseWorksheetCells(rowXml: string, sharedStrings: string[]): string[] {
    const cells = new Map<number, string>();
    const cellRe = /<c\b([^>]*)(?:>([\s\S]*?)<\/c>|\/>)/g;
    let match: RegExpExecArray | null;
    let nextCol = 1;

    while ((match = cellRe.exec(rowXml))) {
      const attrs = this.parseXmlAttributes(match[1]);
      const col = attrs.r ? this.columnNameToNumber(attrs.r.replace(/\d+$/g, "")) : nextCol;
      cells.set(col, this.parseCellValue(match[2] || "", attrs, sharedStrings));
      nextCol = col + 1;
    }

    const maxCol = Math.max(0, ...cells.keys());
    const row: string[] = [];
    for (let col = 1; col <= maxCol; col++) {
      row.push(cells.get(col) || "");
    }
    return row;
  }

  private parseCellValue(innerXml: string, attrs: Record<string, string>, sharedStrings: string[]): string {
    if (attrs.t === "inlineStr") {
      return extractOpenXmlTextValues(innerXml, { namespaceAgnostic: true }).join("");
    }

    const value = this.extractFirstTag(innerXml, "v");
    if (attrs.t === "s") {
      const index = Number(value);
      return Number.isInteger(index) ? sharedStrings[index] || "" : "";
    }
    if (attrs.t === "b") {
      return value === "1" ? "TRUE" : "FALSE";
    }
    if (value !== undefined) {
      return decodeXmlText(value);
    }

    return extractOpenXmlTextValues(innerXml, { namespaceAgnostic: true }).join("");
  }

  private detectWorksheetRange(xml: string, parsedRows: ParsedWorksheetRows): string {
    const dimension = /<dimension\b[^>]*\bref=["']([^"']+)["']/.exec(xml)?.[1];
    if (dimension) return decodeXmlText(dimension);

    const endCol = Math.max(1, parsedRows.maxCol);
    const endRow = Math.max(1, parsedRows.maxRow || parsedRows.totalRows);
    return `A1:${this.toCellAddress(endCol, endRow)}`;
  }

  private extractFirstTag(xml: string, tagName: string): string | undefined {
    const re = new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`);
    const match = re.exec(xml);
    return match ? decodeXmlText(match[1]) : undefined;
  }

  private parseXmlAttributes(tagXml: string): Record<string, string> {
    const attrs: Record<string, string> = {};
    const attrRe = /([\w:-]+)\s*=\s*["']([^"']*)["']/g;
    let match: RegExpExecArray | null;
    while ((match = attrRe.exec(tagXml))) {
      attrs[match[1]] = match[2];
    }
    return attrs;
  }

  private normalizeWorkbookTarget(targetPath: string): string {
    const normalized = targetPath.replace(/\\/g, "/").replace(/^\/+/, "");
    if (normalized.startsWith("xl/")) return path.posix.normalize(normalized);
    return path.posix.normalize(path.posix.join("xl", normalized));
  }

  private sheetPartNumber(partName: string): number {
    return Number(/sheet(\d+)\.xml$/i.exec(partName)?.[1] || 0);
  }

  private slidePartNumber(partName: string): number {
    return Number(/slide(\d+)\.xml$/i.exec(partName)?.[1] || 0);
  }

  private columnNameToNumber(name: string): number {
    return name.toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0);
  }

  private toCellAddress(col: number, row: number): string {
    let value = col;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return `${name || "A"}${row}`;
  }

  private getFileType(filePath: string): KnowledgeFileType {
    const ext = path.extname(filePath).toLowerCase();
    switch (ext) {
      case ".xlsx": return "xlsx";
      case ".xlsm": return "xlsm";
      case ".xlsb": return "xlsb";
      case ".csv": return "csv";
      case ".json": return "json";
      case ".docx": return "docx";
      case ".pptx": return "pptx";
      case ".md": return "md";
      case ".txt": return "txt";
      default: throw new Error(`不支持的文件扩展名: ${ext}`);
    }
  }
}
