import { escapeXmlTextWithQuotes as escapeXml } from "../../../shared/xmlEntities";

export interface SlideInput {
  title: string;
  body: string;
  layout?: string;
}

export function normalizeSlidesParam(params?: Record<string, unknown>): SlideInput[] {
  const rawSlides = params?.slides;
  if (Array.isArray(rawSlides)) {
    return rawSlides.map(normalizeSlideInput).filter((slide): slide is SlideInput => Boolean(slide));
  }

  const title = stringParam(params, "title") || stringParam(params, "heading") || "";
  const body = bodyParam(params);
  if (!title && !body) return [];
  return [{ title, body, layout: stringParam(params, "layout") }];
}

function normalizeSlideInput(value: unknown): SlideInput | undefined {
  if (typeof value === "string") return { title: value, body: "" };
  if (!value || typeof value !== "object" || Array.isArray(value)) return undefined;
  const record = value as Record<string, unknown>;
  const title = stringParam(record, "title") || stringParam(record, "heading") || "";
  const body = bodyParam(record);
  if (!title && !body) return undefined;
  return { title, body, layout: stringParam(record, "layout") };
}

function bodyParam(params?: Record<string, unknown>): string {
  if (!params) return "";
  const body = stringParam(params, "body") || stringParam(params, "content") || stringParam(params, "text");
  if (body) return body;
  const bullets = params.bullets ?? params.items ?? params.points;
  if (Array.isArray(bullets)) {
    return bullets.map((item) => String(item)).filter(Boolean).join("\n");
  }
  return "";
}

export function contentSlideXml(title: string, body: string, layout?: string): string {
  const isBlank = layout === "blank" && !title && !body;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${isBlank ? "" : `${slideTextShapeXml(2, "标题", title, 685800, 457200, 10820400, 777240, 3000)}${slideTextShapeXml(3, "正文", body, 914400, 1508760, 10363200, 4648200, 1800)}`}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function slideTextShapeXml(
  id: number,
  name: string,
  text: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  fontSize: number
): string {
  const paragraphs = text
    ? text.split(/\r?\n/).map((line) => textParagraphXml(line, fontSize)).join("")
    : textParagraphXml("", fontSize);
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody></p:sp>`;
}

function textParagraphXml(text: string, fontSize: number): string {
  return `<a:p><a:r><a:rPr lang="zh-CN" sz="${fontSize}"/><a:t>${escapeXml(text)}</a:t></a:r><a:endParaRPr lang="zh-CN" sz="${fontSize}"/></a:p>`;
}

export function emptySlideRelsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`;
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
}
