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
import {
  detectWorksheetRange,
  parseWorksheetRows,
  readSharedStrings,
  readWorkbookSheets,
} from "./excelWorkbookParser";
import { flattenJson } from "./jsonFlatten";
import type { KnowledgeFileType } from "./types";

const MAX_EXCEL_PARSE_BYTES = 25 * 1024 * 1024;
const MAX_EXCEL_DATA_ROWS = 500;
const MAX_JSON_LINES = 2000;

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
    const sharedStrings = await readSharedStrings(zip);
    const sheets = await readWorkbookSheets(zip);
    const chunks: RawChunk[] = [];

    for (const sheetInfo of sheets) {
      const sheetPart = zip.file(sheetInfo.partName);
      if (!sheetPart) continue;

      const xml = await sheetPart.async("text");
      const parsedRows = parseWorksheetRows(xml, sharedStrings, MAX_EXCEL_DATA_ROWS + 1);
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
          tableRange: detectWorksheetRange(xml, parsedRows),
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

  private slidePartNumber(partName: string): number {
    return Number(/slide(\d+)\.xml$/i.exec(partName)?.[1] || 0);
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
