import * as path from "path";

import JSZip from "jszip";
import { extractOpenXmlTextValues } from "../shared/openXmlText";
import { decodeXmlText } from "../shared/xmlEntities";

export interface WorkbookSheetInfo {
  name: string;
  partName: string;
}

export interface ParsedWorksheetRows {
  rows: string[][];
  totalRows: number;
  maxCol: number;
  maxRow: number;
}

export async function readWorkbookSheets(zip: JSZip): Promise<WorkbookSheetInfo[]> {
  const workbookPart = zip.file("xl/workbook.xml");
  const relsPart = zip.file("xl/_rels/workbook.xml.rels");
  if (!workbookPart || !relsPart) return fallbackWorksheetParts(zip);

  const workbookXml = await workbookPart.async("text");
  const relsXml = await relsPart.async("text");
  const relationshipTargets = parseWorkbookRelationships(relsXml);
  const sheets: WorkbookSheetInfo[] = [];
  const sheetRe = /<sheet\b[^>]*\/?>/g;
  let match: RegExpExecArray | null;

  while ((match = sheetRe.exec(workbookXml))) {
    const attrs = parseXmlAttributes(match[0]);
    const relId = attrs["r:id"] || attrs.id;
    const target = relId ? relationshipTargets.get(relId) : undefined;
    if (!target) continue;
    sheets.push({
      name: decodeXmlText(attrs.name || `Sheet${sheets.length + 1}`),
      partName: target,
    });
  }

  return sheets.length > 0 ? sheets : fallbackWorksheetParts(zip);
}

export async function readSharedStrings(zip: JSZip): Promise<string[]> {
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

export function parseWorksheetRows(
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
    const attrs = parseXmlAttributes(rowMatch[1]);
    const rowNumber = Number(attrs.r) || totalRows;
    maxRow = Math.max(maxRow, rowNumber);

    const parsed = parseWorksheetCells(rowMatch[2], sharedStrings);
    maxCol = Math.max(maxCol, parsed.length);
    if (rows.length < rowLimit) {
      rows.push(parsed);
    }
  }

  return { rows, totalRows, maxCol, maxRow };
}

export function detectWorksheetRange(xml: string, parsedRows: ParsedWorksheetRows): string {
  const dimension = /<dimension\b[^>]*\bref=["']([^"']+)["']/.exec(xml)?.[1];
  if (dimension) return decodeXmlText(dimension);

  const endCol = Math.max(1, parsedRows.maxCol);
  const endRow = Math.max(1, parsedRows.maxRow || parsedRows.totalRows);
  return `A1:${toCellAddress(endCol, endRow)}`;
}

function parseWorkbookRelationships(relsXml: string): Map<string, string> {
  const targets = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/?>/g;
  let match: RegExpExecArray | null;
  while ((match = relRe.exec(relsXml))) {
    const attrs = parseXmlAttributes(match[0]);
    if (attrs.Id && attrs.Target) {
      targets.set(attrs.Id, normalizeWorkbookTarget(attrs.Target));
    }
  }
  return targets;
}

function fallbackWorksheetParts(zip: JSZip): WorkbookSheetInfo[] {
  return Object.keys(zip.files)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((a, b) => sheetPartNumber(a) - sheetPartNumber(b))
    .map((partName, index) => ({ name: `Sheet${index + 1}`, partName }));
}

function parseWorksheetCells(rowXml: string, sharedStrings: string[]): string[] {
  const cells = new Map<number, string>();
  const cellRe = /<c\b([^>]*)(?:>([\s\S]*?)<\/c>|\/>)/g;
  let match: RegExpExecArray | null;
  let nextCol = 1;

  while ((match = cellRe.exec(rowXml))) {
    const attrs = parseXmlAttributes(match[1]);
    const col = attrs.r ? columnNameToNumber(attrs.r.replace(/\d+$/g, "")) : nextCol;
    cells.set(col, parseCellValue(match[2] || "", attrs, sharedStrings));
    nextCol = col + 1;
  }

  const maxCol = Math.max(0, ...cells.keys());
  const row: string[] = [];
  for (let col = 1; col <= maxCol; col++) {
    row.push(cells.get(col) || "");
  }
  return row;
}

function parseCellValue(innerXml: string, attrs: Record<string, string>, sharedStrings: string[]): string {
  if (attrs.t === "inlineStr") {
    return extractOpenXmlTextValues(innerXml, { namespaceAgnostic: true }).join("");
  }

  const value = extractFirstTag(innerXml, "v");
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

function extractFirstTag(xml: string, tagName: string): string | undefined {
  const re = new RegExp(`<${tagName}\\b[^>]*>([\\s\\S]*?)<\\/${tagName}>`);
  const match = re.exec(xml);
  return match ? decodeXmlText(match[1]) : undefined;
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
  if (normalized.startsWith("xl/")) return path.posix.normalize(normalized);
  return path.posix.normalize(path.posix.join("xl", normalized));
}

function sheetPartNumber(partName: string): number {
  return Number(/sheet(\d+)\.xml$/i.exec(partName)?.[1] || 0);
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
  return `${name || "A"}${row}`;
}
