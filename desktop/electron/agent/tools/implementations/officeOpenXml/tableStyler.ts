/**
 * Open XML 表格样式器
 *
 * 关联模块：
 * - tableInspector.ts: 可先检查表格结构。
 * - officeOpenXmlFileBridge.ts: 将样式能力暴露给工具执行器。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import type {
  OfficeOpenXmlDocumentType,
  OfficeOpenXmlTableStyleInput,
  OfficeOpenXmlTableStylePreset,
  OfficeOpenXmlTableStyleResult,
} from "./types";

const WORD_DOCUMENT_PART = "word/document.xml";
const PRESENTATION_SLIDE_RE = /^ppt\/slides\/slide\d+\.xml$/;
const SPREADSHEET_SHEET_RE = /^xl\/worksheets\/sheet\d+\.xml$/;

function detectDocumentType(filePath: string): OfficeOpenXmlDocumentType {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".docx") return "word";
  if (ext === ".pptx") return "presentation";
  if (ext === ".xlsx") return "spreadsheet";
  throw new Error(`仅支持 .docx、.pptx 和 .xlsx 文件: ${filePath}`);
}

function defaultOutputPath(filePath: string): string {
  const dir = path.dirname(filePath);
  const ext = path.extname(filePath);
  const base = path.basename(filePath, ext);
  return path.join(dir, `${base}-styled${ext}`);
}

function styleColor(style: OfficeOpenXmlTableStylePreset): string {
  if (style === "financial") return "385723";
  if (style === "compact") return "5B9BD5";
  return "1F4E79";
}

function styleWordHeaderRow(rowXml: string, color: string): string {
  return rowXml
    .replace(/<w:tc\b[^>]*>[\s\S]*?<\/w:tc>/g, (cellXml) => {
      if (/<w:tcPr\b/.test(cellXml)) {
        return cellXml.replace(/<w:tcPr\b[^>]*>/, (tag) => `${tag}<w:shd w:fill="${color}" />`);
      }
      return cellXml.replace(/<w:tc\b[^>]*>/, (tag) => `${tag}<w:tcPr><w:shd w:fill="${color}" /></w:tcPr>`);
    })
    .replace(/<w:r\b([^>]*)>/g, '<w:r$1><w:rPr><w:b /></w:rPr>');
}

function styleWordTables(xml: string, color: string): { xml: string; changed: boolean } {
  let changed = false;
  const nextXml = xml.replace(/<w:tbl\b[\s\S]*?<\/w:tbl>/g, (tableXml) => {
    const styledTable = tableXml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/, (rowXml) => {
      changed = true;
      return styleWordHeaderRow(rowXml, color);
    });
    return styledTable;
  });
  return { xml: nextXml, changed };
}

function stylePresentationHeaderRow(rowXml: string, color: string): string {
  return rowXml
    .replace(/<a:tc\b[^>]*>[\s\S]*?<\/a:tc>/g, (cellXml) => {
      if (/<a:tcPr\b/.test(cellXml)) {
        return cellXml.replace(/<a:tcPr\b[^>]*>/, (tag) => `${tag}<a:solidFill><a:srgbClr val="${color}" /></a:solidFill>`);
      }
      return cellXml.replace(
        /<\/a:tc>/,
        `<a:tcPr><a:solidFill><a:srgbClr val="${color}" /></a:solidFill></a:tcPr></a:tc>`
      );
    })
    .replace(/<a:r\b([^>]*)>/g, '<a:r$1><a:rPr b="1" />');
}

function stylePresentationTables(xml: string, color: string): { xml: string; changed: boolean } {
  let changed = false;
  const nextXml = xml.replace(/<a:tbl\b[\s\S]*?<\/a:tbl>/g, (tableXml) => {
    const styledTable = tableXml.replace(/<a:tr\b[\s\S]*?<\/a:tr>/, (rowXml) => {
      changed = true;
      return stylePresentationHeaderRow(rowXml, color);
    });
    return styledTable;
  });
  return { xml: nextXml, changed };
}

function styleSpreadsheetHeaderRow(rowXml: string, styleIndex: number): string {
  return rowXml.replace(/<c\b([^>]*)>/g, (tag, attributes: string) => {
    if (!/\br="[^"]+"/.test(attributes)) return tag;
    const withoutStyle = attributes.replace(/\s+s="[^"]*"/g, "");
    return `<c${withoutStyle.replace(/\br="[^"]+"/, (cellRef) => `${cellRef} s="${styleIndex}"`)}>`;
  });
}

function styleSpreadsheetRows(xml: string, styleIndex: number): { xml: string; changed: boolean } {
  let changed = false;
  const nextXml = xml.replace(/<row\b[\s\S]*?<\/row>/, (rowXml) => {
    changed = true;
    return styleSpreadsheetHeaderRow(rowXml, styleIndex);
  });
  return { xml: nextXml, changed };
}

function spreadsheetStylesXml(color: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font /><font><b /></font></fonts>
  <fills count="3">
    <fill><patternFill patternType="none" /></fill>
    <fill><patternFill patternType="gray125" /></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF${color}" /><bgColor indexed="64" /></patternFill></fill>
  </fills>
  <borders count="1"><border><left /><right /><top /><bottom /><diagonal /></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" />
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0" /></cellStyles>
</styleSheet>`;
}

type StyleCollectionName = "fonts" | "fills" | "cellXfs";
type StyleItemName = "font" | "fill" | "xf";

const STYLE_COLLECTION_ORDER = [
  "numFmts",
  "fonts",
  "fills",
  "borders",
  "cellStyleXfs",
  "cellXfs",
  "cellStyles",
  "dxfs",
  "tableStyles",
  "colors",
  "extLst",
] as const;

interface AppendedStyleCollection {
  xml: string;
  index: number;
}

function updateCollectionCount(openingTag: string, count: number): string {
  if (/\bcount="[^"]*"/.test(openingTag)) {
    return openingTag.replace(/\bcount="[^"]*"/, `count="${count}"`);
  }
  return openingTag.replace(/>$/, ` count="${count}">`);
}

function findXmlTagEnd(xml: string, tagStart: number): number {
  let quote: '"' | "'" | undefined;
  for (let index = tagStart + 1; index < xml.length; index += 1) {
    const char = xml[index];
    if (quote) {
      if (char === quote) quote = undefined;
      continue;
    }
    if (char === '"' || char === "'") {
      quote = char;
      continue;
    }
    if (char === ">") return index;
  }
  return -1;
}

function findStyleElementBounds(
  xml: string,
  elementName: string
): { start: number; end: number } | undefined {
  const rootMatch = /<styleSheet\b[^>]*>/.exec(xml);
  const rootEnd = xml.lastIndexOf("</styleSheet>");
  if (!rootMatch || rootEnd < rootMatch.index + rootMatch[0].length) return undefined;

  let cursor = rootMatch.index + rootMatch[0].length;
  let depth = 0;
  let directChildStart = -1;
  let directChildName = "";

  while (cursor < rootEnd) {
    const tagStart = xml.indexOf("<", cursor);
    if (tagStart < 0 || tagStart >= rootEnd) return undefined;
    if (xml.startsWith("<!--", tagStart)) {
      const commentEnd = xml.indexOf("-->", tagStart + 4);
      if (commentEnd < 0) return undefined;
      cursor = commentEnd + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", tagStart)) {
      const cdataEnd = xml.indexOf("]]>", tagStart + 9);
      if (cdataEnd < 0) return undefined;
      cursor = cdataEnd + 3;
      continue;
    }
    if (xml.startsWith("<?", tagStart)) {
      const processingEnd = xml.indexOf("?>", tagStart + 2);
      if (processingEnd < 0) return undefined;
      cursor = processingEnd + 2;
      continue;
    }

    const tagEnd = findXmlTagEnd(xml, tagStart);
    if (tagEnd < 0 || tagEnd >= rootEnd) return undefined;
    const tag = xml.slice(tagStart, tagEnd + 1);
    if (/^<!/.test(tag)) {
      cursor = tagEnd + 1;
      continue;
    }

    const name = /^<\s*\/?\s*([A-Za-z_][\w:.-]*)/.exec(tag)?.[1];
    if (!name) {
      cursor = tagEnd + 1;
      continue;
    }
    const isClosing = /^<\s*\//.test(tag);
    const isSelfClosing = /\/\s*>$/.test(tag);

    if (isClosing) {
      if (depth > 0) {
        depth -= 1;
        if (depth === 0) {
          if (directChildName === elementName) {
            return { start: directChildStart, end: tagEnd + 1 };
          }
          directChildStart = -1;
          directChildName = "";
        }
      }
    } else if (depth === 0) {
      if (isSelfClosing) {
        if (name === elementName) {
          return { start: tagStart, end: tagEnd + 1 };
        }
      } else {
        directChildStart = tagStart;
        directChildName = name;
        depth = 1;
      }
    } else if (!isSelfClosing) {
      depth += 1;
    }

    cursor = tagEnd + 1;
  }
  return undefined;
}

function insertStyleCollection(
  xml: string,
  collectionName: StyleCollectionName,
  collectionXml: string
): string {
  const collectionIndex = STYLE_COLLECTION_ORDER.indexOf(collectionName);
  for (let index = collectionIndex + 1; index < STYLE_COLLECTION_ORDER.length; index += 1) {
    const next = findStyleElementBounds(xml, STYLE_COLLECTION_ORDER[index]);
    if (next) {
      return `${xml.slice(0, next.start)}${collectionXml}${xml.slice(next.start)}`;
    }
  }
  for (let index = collectionIndex - 1; index >= 0; index -= 1) {
    const previous = findStyleElementBounds(xml, STYLE_COLLECTION_ORDER[index]);
    if (previous) {
      return `${xml.slice(0, previous.end)}${collectionXml}${xml.slice(previous.end)}`;
    }
  }
  if (/<styleSheet\b[^>]*>/.test(xml)) {
    return xml.replace(/<styleSheet\b[^>]*>/, (tag) => `${tag}${collectionXml}`);
  }
  throw new Error("Excel 样式表缺少 styleSheet 根节点");
}

function appendStyleCollectionItem(
  xml: string,
  collectionName: StyleCollectionName,
  itemName: StyleItemName,
  itemXml: string
): AppendedStyleCollection {
  const collectionRe = new RegExp(
    `<${collectionName}\\b[^>]*>([\\s\\S]*?)<\\/${collectionName}>`
  );
  const collectionMatch = collectionRe.exec(xml);
  if (collectionMatch) {
    const collectionXml = collectionMatch[0];
    const openingTag = new RegExp(`^<${collectionName}\\b[^>]*>`).exec(collectionXml)?.[0];
    if (!openingTag) {
      throw new Error(`无法读取 Excel 样式集合: ${collectionName}`);
    }
    const itemCount = (collectionMatch[1].match(new RegExp(`<${itemName}\\b`, "g")) || []).length;
    const nextCollection = collectionXml
      .replace(openingTag, updateCollectionCount(openingTag, itemCount + 1))
      .replace(`</${collectionName}>`, `${itemXml}</${collectionName}>`);
    return {
      xml: xml.replace(collectionXml, nextCollection),
      index: itemCount,
    };
  }

  const selfClosingRe = new RegExp(`<${collectionName}\\b[^>]*/>`);
  const selfClosingMatch = selfClosingRe.exec(xml);
  if (selfClosingMatch) {
    const openingTag = updateCollectionCount(
      selfClosingMatch[0].replace(/\s*\/>$/, ">"),
      1
    );
    return {
      xml: xml.replace(
        selfClosingMatch[0],
        `${openingTag}${itemXml}</${collectionName}>`
      ),
      index: 0,
    };
  }

  const collectionXml = `<${collectionName} count="1">${itemXml}</${collectionName}>`;
  return {
    xml: insertStyleCollection(xml, collectionName, collectionXml),
    index: 0,
  };
}

function appendSpreadsheetHeaderStyle(xml: string, color: string): { xml: string; styleIndex: number } {
  const font = appendStyleCollectionItem(
    xml,
    "fonts",
    "font",
    "<font><b /></font>"
  );
  const fill = appendStyleCollectionItem(
    font.xml,
    "fills",
    "fill",
    `<fill><patternFill patternType="solid"><fgColor rgb="FF${color}" /><bgColor indexed="64" /></patternFill></fill>`
  );
  const cellXf = appendStyleCollectionItem(
    fill.xml,
    "cellXfs",
    "xf",
    `<xf numFmtId="0" fontId="${font.index}" fillId="${fill.index}" borderId="0" xfId="0" applyFont="1" applyFill="1" />`
  );
  return {
    xml: cellXf.xml,
    styleIndex: cellXf.index,
  };
}

async function ensureSpreadsheetStyleParts(
  zip: JSZip,
  color: string,
  changedParts: string[]
): Promise<number> {
  const existingStyles = zip.file("xl/styles.xml");
  let styleIndex = 1;
  if (existingStyles) {
    const appended = appendSpreadsheetHeaderStyle(await existingStyles.async("text"), color);
    zip.file("xl/styles.xml", appended.xml);
    styleIndex = appended.styleIndex;
  } else {
    zip.file("xl/styles.xml", spreadsheetStylesXml(color));
  }
  changedParts.push("xl/styles.xml");

  const contentTypes = zip.file("[Content_Types].xml");
  if (contentTypes) {
    const xml = await contentTypes.async("text");
    if (!xml.includes('PartName="/xl/styles.xml"')) {
      zip.file(
        "[Content_Types].xml",
        xml.replace(
          "</Types>",
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" /></Types>'
        )
      );
      changedParts.push("[Content_Types].xml");
    }
  }

  const rels = zip.file("xl/_rels/workbook.xml.rels");
  if (rels) {
    const xml = await rels.async("text");
    if (!xml.includes("/styles")) {
      zip.file(
        "xl/_rels/workbook.xml.rels",
        xml.replace(
          "</Relationships>",
          '<Relationship Id="rIdOpenXmlStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" /></Relationships>'
        )
      );
      changedParts.push("xl/_rels/workbook.xml.rels");
    }
  }
  return styleIndex;
}

export async function applyOfficeOpenXmlTableStyle(
  input: OfficeOpenXmlTableStyleInput
): Promise<OfficeOpenXmlTableStyleResult> {
  const documentType = detectDocumentType(input.filePath);
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const changedParts: string[] = [];
  const color = styleColor(input.style);

  if (documentType === "word") {
    const part = zip.file(WORD_DOCUMENT_PART);
    if (part) {
      const styled = styleWordTables(await part.async("text"), color);
      if (styled.changed) {
        zip.file(WORD_DOCUMENT_PART, styled.xml);
        changedParts.push(WORD_DOCUMENT_PART);
      }
    }
  }

  if (documentType === "presentation") {
    const partNames = Object.keys(zip.files).filter((name) => PRESENTATION_SLIDE_RE.test(name)).sort();
    for (const partName of partNames) {
      const part = zip.file(partName);
      if (!part) continue;
      const styled = stylePresentationTables(await part.async("text"), color);
      if (styled.changed) {
        zip.file(partName, styled.xml);
        changedParts.push(partName);
      }
    }
  }

  if (documentType === "spreadsheet") {
    const partNames = Object.keys(zip.files).filter((name) => SPREADSHEET_SHEET_RE.test(name)).sort();
    const styleIndex = partNames.length > 0
      ? await ensureSpreadsheetStyleParts(zip, color, changedParts)
      : 1;
    for (const partName of partNames) {
      const part = zip.file(partName);
      if (!part) continue;
      const styled = styleSpreadsheetRows(await part.async("text"), styleIndex);
      if (styled.changed) {
        zip.file(partName, styled.xml);
        changedParts.push(partName);
      }
    }
  }

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));

  return {
    engine: "openxml",
    operation: "applyTableStyle",
    documentType,
    filePath: input.filePath,
    outputPath,
    target: input.target,
    style: input.style,
    changedParts,
  };
}
