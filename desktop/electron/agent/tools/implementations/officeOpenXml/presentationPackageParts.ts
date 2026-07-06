import type JSZip from "jszip";

export const CONTENT_TYPES_PART = "[Content_Types].xml";

const SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";

export interface SlideEntry {
  index: number;
  relId: string;
  partName?: string;
}

export function collectSlideEntries(presentationXml: string, relsXml: string): SlideEntry[] {
  const relationshipTargets = new Map<string, string>();
  for (const match of relsXml.matchAll(/<Relationship\b[^>]*(?:\/>|><\/Relationship>)/g)) {
    const relId = getXmlAttr(match[0], "Id");
    const target = getXmlAttr(match[0], "Target");
    if (relId && target) relationshipTargets.set(relId, normalizeSlidePartName(target));
  }

  const entries: SlideEntry[] = [];
  for (const match of presentationXml.matchAll(/<p:sldId\b[^>]*(?:\/>|><\/p:sldId>)/g)) {
    const relId = getXmlAttr(match[0], "r:id");
    if (!relId) continue;
    entries.push({
      index: entries.length + 1,
      relId,
      partName: relationshipTargets.get(relId),
    });
  }
  return entries;
}

export function nextSlidePartNumber(zip: JSZip): number {
  const used = Object.keys(zip.files)
    .map((name) => /^ppt\/slides\/slide(\d+)\.xml$/.exec(name)?.[1])
    .filter((value): value is string => Boolean(value))
    .map(Number);
  return Math.max(0, ...used) + 1;
}

export function nextPresentationSlideId(presentationXml: string): number {
  const ids = [...presentationXml.matchAll(/<p:sldId\b[^>]*\bid=["'](\d+)["']/g)]
    .map((match) => Number(match[1]))
    .filter((value) => Number.isFinite(value));
  return Math.max(255, ...ids) + 1;
}

export function nextRelationshipNumber(relsXml: string): number {
  const ids = [...relsXml.matchAll(/\bId=["']rId(\d+)["']/g)]
    .map((match) => Number(match[1]))
    .filter((value) => Number.isFinite(value));
  return Math.max(0, ...ids) + 1;
}

export function insertSlideId(presentationXml: string, slideId: number, relId: string): string {
  const slideXml = `<p:sldId id="${slideId}" r:id="${relId}"/>`;
  if (presentationXml.includes("</p:sldIdLst>")) {
    return presentationXml.replace("</p:sldIdLst>", `${slideXml}</p:sldIdLst>`);
  }
  const slideListXml = `<p:sldIdLst>${slideXml}</p:sldIdLst>`;
  return presentationXml.includes("<p:sldSz")
    ? presentationXml.replace("<p:sldSz", `${slideListXml}<p:sldSz`)
    : presentationXml.replace("</p:presentation>", `${slideListXml}</p:presentation>`);
}

export function insertPresentationRelationship(relsXml: string, relId: string, target: string): string {
  const relXml = `<Relationship Id="${relId}" Type="${SLIDE_REL_TYPE}" Target="${target}"/>`;
  return relsXml.includes("</Relationships>")
    ? relsXml.replace("</Relationships>", `${relXml}</Relationships>`)
    : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relXml}</Relationships>`;
}

export async function contentTypesXmlOrDefault(zip: JSZip): Promise<string> {
  const part = zip.file(CONTENT_TYPES_PART);
  if (part) return part.async("text");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`;
}

export function ensureSlideContentType(contentTypesXml: string, slidePartName: string): string {
  const partName = `/${slidePartName}`;
  if (contentTypesXml.includes(`PartName="${partName}"`) || contentTypesXml.includes(`PartName='${partName}'`)) {
    return contentTypesXml;
  }
  const overrideXml = `<Override PartName="${partName}" ContentType="${SLIDE_CONTENT_TYPE}"/>`;
  return contentTypesXml.includes("</Types>")
    ? contentTypesXml.replace("</Types>", `${overrideXml}</Types>`)
    : `${contentTypesXml}${overrideXml}`;
}

export function getXmlAttr(xml: string, attrName: string): string | undefined {
  const escapedName = attrName.replace(/:/g, "\\:");
  const match = xml.match(new RegExp(`\\b${escapedName}=["']([^"']+)["']`));
  return match?.[1];
}

function normalizeSlidePartName(target: string): string {
  const normalized = target.replace(/^\/+/, "");
  return normalized.startsWith("ppt/") ? normalized : `ppt/${normalized}`;
}
