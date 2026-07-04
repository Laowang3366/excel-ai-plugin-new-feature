import type JSZip from "jszip";
import { decodeXmlText } from "./xmlEntities";

export interface OpenXmlTextExtractOptions {
  tagName?: string;
  namespaceAgnostic?: boolean;
  includeEmpty?: boolean;
}

export interface OpenXmlTextPart {
  partName: string;
  text: string;
  textLength: number;
}

export function extractOpenXmlTextValues(
  xml: string,
  options: OpenXmlTextExtractOptions = {}
): string[] {
  const tagPattern = buildTagPattern(options.tagName || "t", options.namespaceAgnostic ?? false);
  const re = new RegExp(`<${tagPattern}\\b[^>]*>([\\s\\S]*?)<\\/${tagPattern}>`, "g");
  const values = Array.from(xml.matchAll(re), (match) => decodeXmlText(match[1]));
  if (options.includeEmpty === false) {
    return values.filter((value) => value.trim().length > 0);
  }
  return values;
}

export function extractOpenXmlText(
  xml: string,
  options: OpenXmlTextExtractOptions = {},
  separator = "\n"
): string {
  return extractOpenXmlTextValues(xml, options).join(separator);
}

export function extractOpenXmlParagraphTexts(xml: string): string[] {
  const lines: string[] = [];
  const paragraphRe = /<(?:\w+:)?p\b[^>]*>([\s\S]*?)<\/(?:\w+:)?p>/g;
  let match: RegExpExecArray | null;
  while ((match = paragraphRe.exec(xml))) {
    const text = extractOpenXmlTextValues(match[1], { namespaceAgnostic: true })
      .join("")
      .replace(/\s+/g, " ")
      .trim();
    if (text) lines.push(text);
  }
  if (lines.length > 0) return lines;
  return extractOpenXmlTextValues(xml, { namespaceAgnostic: true })
    .map((value) => value.trim())
    .filter(Boolean);
}

export async function readOpenXmlTextParts(
  zip: JSZip,
  partRe: RegExp,
  options: OpenXmlTextExtractOptions = {}
): Promise<OpenXmlTextPart[]> {
  const parts: OpenXmlTextPart[] = [];
  const partNames = Object.keys(zip.files).filter((name) => partRe.test(name)).sort();
  for (const partName of partNames) {
    const part = zip.file(partName);
    if (!part) continue;
    const text = extractOpenXmlText(await part.async("text"), options);
    if (!text) continue;
    parts.push({ partName, text, textLength: text.length });
  }
  return parts;
}

function buildTagPattern(tagName: string, namespaceAgnostic: boolean): string {
  const localName = tagName.includes(":") ? tagName.split(":").pop() || tagName : tagName;
  if (namespaceAgnostic) return `(?:\\w+:)?${escapeRegExp(localName)}`;
  return escapeRegExp(tagName);
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
