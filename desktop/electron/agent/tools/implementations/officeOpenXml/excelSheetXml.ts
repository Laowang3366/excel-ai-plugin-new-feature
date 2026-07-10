import JSZip from "jszip";
import {
  escapeXmlAttribute as escapeXml,
  escapeXmlTextWithQuotes as escapeXmlText,
} from "../../../shared/xmlEntities";
import { formulaCellXml } from "./excelFormulaXml";

export function addWorkbookBaseParts(zip: JSZip, sheetNames: string[]): void {
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  ${sheetNames.map((_name, index) => `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("\n  ")}
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetNames.map((name, index) => `<sheet name="${escapeXml(name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join("\n    ")}
  </sheets>
</workbook>`);
  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetNames.map((_name, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`).join("\n  ")}
  <Relationship Id="rId${sheetNames.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`);
}

export function worksheetXml(sheetData: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${sheetData}</sheetData>
</worksheet>`;
}

export function buildSheetDataXml(
  startCell: string,
  values: unknown[][],
  targetRef?: string,
  clearBlankCells = false
): string {
  const start = parseCellAddress(startCell);
  return values.map((row, rowOffset) => {
    const rowNumber = start.row + rowOffset;
    const cells = row
      .map((value, colOffset) => cellXml(toCellAddress(start.col + colOffset, rowNumber), value, targetRef, clearBlankCells))
      .filter(Boolean)
      .join("");
    return `<row r="${rowNumber}">${cells}</row>`;
  }).join("");
}

export function mergeSheetDataXml(
  xml: string,
  startCell: string,
  values: unknown[][],
  targetRef?: string,
  clearBlankCells = false
): string {
  const rows = new Map<number, SheetRow>();
  const sheetDataMatch = /<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml);
  if (sheetDataMatch) {
    const rowRe = /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi;
    let rowMatch: RegExpExecArray | null;
    while ((rowMatch = rowRe.exec(sheetDataMatch[1]))) {
      const parsedRow = parseSheetRow(rowMatch[0]);
      if (parsedRow) rows.set(parsedRow.rowNumber, parsedRow);
    }
  }

  const start = parseCellAddress(startCell);
  values.forEach((row, rowOffset) => {
    const rowNumber = start.row + rowOffset;
    const sheetRow = getOrCreateRow(rows, rowNumber);
    row.forEach((value, colOffset) => {
      const colNumber = start.col + colOffset;
      const nextCellXml = cellXml(toCellAddress(colNumber, rowNumber), value, targetRef, clearBlankCells);
      if (nextCellXml) {
        sheetRow.cells.set(colNumber, nextCellXml);
      } else {
        sheetRow.cells.delete(colNumber);
      }
    });
  });

  const mergedSheetData = [...rows.values()]
    .sort((a, b) => a.rowNumber - b.rowNumber)
    .map((row) => {
      const cellXmls = [...row.cells.entries()]
        .sort(([a], [b]) => a - b)
        .map(([, cell]) => cell)
        .join("");
      return `${row.openingTag}${cellXmls}${row.nonCellXml}</row>`;
    })
    .join("");
  return replaceSheetData(xml, mergedSheetData);
}

interface SheetRow {
  rowNumber: number;
  openingTag: string;
  cells: Map<number, string>;
  nonCellXml: string;
}

function parseSheetRow(rowXml: string): SheetRow | null {
  const openingTag = /^<row\b[^>]*>/.exec(rowXml)?.[0];
  const selfClosingTag = /^<row\b[^>]*\/>/.exec(rowXml)?.[0];
  const sourceTag = openingTag || selfClosingTag;
  if (!sourceTag) return null;

  const rowNumberMatch = /\br="(\d+)"/i.exec(sourceTag);
  if (!rowNumberMatch) return null;

  const body = selfClosingTag
    ? ""
    : rowXml.slice(sourceTag.length, rowXml.length - "</row>".length);
  const cells = new Map<number, string>();
  const cellRe = /<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:>[\s\S]*?<\/c>|\/>)/gi;
  let cellMatch: RegExpExecArray | null;
  while ((cellMatch = cellRe.exec(body))) {
    const parsed = parseCellAddress(cellMatch[1]);
    cells.set(parsed.col, cellMatch[0]);
  }

  return {
    rowNumber: Number(rowNumberMatch[1]),
    openingTag: selfClosingTag ? selfClosingTag.replace(/\/>$/, ">") : sourceTag,
    cells,
    nonCellXml: body.replace(cellRe, ""),
  };
}

function replaceSheetData(xml: string, sheetData: string): string {
  if (/<sheetData\b[\s\S]*?<\/sheetData>/.test(xml)) {
    return xml.replace(/<sheetData\b[^>]*>[\s\S]*?<\/sheetData>/, `<sheetData>${sheetData}</sheetData>`);
  }
  return insertBeforeWorksheetEnd(xml, `<sheetData>${sheetData}</sheetData>`);
}

function insertBeforeWorksheetEnd(xml: string, addition: string): string {
  return xml.includes("</worksheet>") ? xml.replace("</worksheet>", `${addition}</worksheet>`) : `${xml}${addition}`;
}

function getOrCreateRow(rows: Map<number, SheetRow>, rowNumber: number): SheetRow {
  const existing = rows.get(rowNumber);
  if (existing) return existing;
  const next: SheetRow = {
    rowNumber,
    openingTag: `<row r="${rowNumber}">`,
    cells: new Map<number, string>(),
    nonCellXml: "",
  };
  rows.set(rowNumber, next);
  return next;
}

function cellXml(address: string, value: unknown, targetRef?: string, clearBlankCells = false): string {
  if (value === null || value === undefined) return clearBlankCells ? "" : `<c r="${address}"/>`;
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${address}"><v>${value}</v></c>`;
  }
  if (typeof value === "boolean") {
    return `<c r="${address}" t="b"><v>${value ? 1 : 0}</v></c>`;
  }
  const text = String(value);
  if (clearBlankCells && text === "") return "";
  if (text.startsWith("=")) {
    return formulaCellXml(address, text.slice(1), targetRef);
  }
  return `<c r="${address}" t="inlineStr"><is><t>${escapeXmlText(text)}</t></is></c>`;
}

function parseCellAddress(address: string): { col: number; row: number } {
  const match = /^([A-Z]+)(\d+)$/i.exec(address.trim());
  if (!match) return { col: 1, row: 1 };
  return { col: columnNameToNumber(match[1]), row: Number(match[2]) };
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
  return `${name}${row}`;
}
