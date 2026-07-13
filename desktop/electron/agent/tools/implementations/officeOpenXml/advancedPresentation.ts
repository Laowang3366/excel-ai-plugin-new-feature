/**
 * PowerPoint Open XML 高级操作
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: 将统一 Office action 路由到本模块。
 * - tableStyler.ts: 表格美化仍由既有通用表格样式器处理。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import pptxgen from "pptxgenjs";
import { failedResult, needsComResult, unsupportedResult } from "../../officeCore/results";
import type { OfficeActionKind, OfficeActionResult } from "../../officeCore/types";
import {
  collectSlideEntries,
  CONTENT_TYPES_PART,
  contentTypesXmlOrDefault,
  ensureSlideContentType,
  getXmlAttr,
  insertPresentationRelationship,
  insertSlideId,
  nextPresentationSlideId,
  nextRelationshipNumber,
  nextSlidePartNumber,
} from "./presentationPackageParts";
import { contentSlideXml, emptySlideRelsXml, normalizeSlidesParam } from "./presentationSlideContent";
import { createOpenXmlDoneResult } from "./actionResult";

export interface PresentationAdvancedActionInput {
  operation: string;
  filePath: string;
  outputPath?: string;
  target?: string;
  action?: OfficeActionKind;
  params?: Record<string, unknown>;
}

const PRESENTATION_SLIDE_RE = /^ppt\/slides\/slide\d+\.xml$/;
const PRESENTATION_PART = "ppt/presentation.xml";
const PRESENTATION_RELS_PART = "ppt/_rels/presentation.xml.rels";
const ADD_SLIDE_OPERATIONS = new Set(["addSlide", "addSlides", "appendSlide", "appendSlides", "addSlideContent"]);
const presentationDone = createOpenXmlDoneResult("presentation");

export async function applyPresentationAdvancedAction(
  input: PresentationAdvancedActionInput
): Promise<OfficeActionResult> {
  try {
    if (input.operation === "createPresentation") {
      return await createPresentation(input);
    }
    if (input.operation === "applyTheme") {
      return await applyTheme(input);
    }
    if (input.operation === "deleteSlides") {
      return await deleteSlides(input);
    }
    if (ADD_SLIDE_OPERATIONS.has(input.operation)) {
      return await addSlides(input);
    }
    if (
      input.operation === "normalizeLayouts" ||
      input.operation === "alignShapes" ||
      input.operation === "insertChart" ||
      input.operation === "replacePictureSlot"
    ) {
      return needsComResult({
        app: "presentation",
        action: input.action || "edit",
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: `${input.operation} 需要完整坐标、关系图或媒体包维护，需显式 COM 兜底`,
      });
    }

    return unsupportedResult({
      app: "presentation",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: `暂不支持 PPT Open XML 高级操作: ${input.operation}`,
    });
  } catch (error) {
    return failedResult({
      app: "presentation",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
    }, error);
  }
}

async function createPresentation(input: PresentationAdvancedActionInput): Promise<OfficeActionResult> {
  const outputPath = input.outputPath || input.filePath;
  const title = stringParam(input.params, "title") || "新建演示文稿";
  const subtitle = stringParam(input.params, "subtitle") || "";
  const presentation = new pptxgen();
  presentation.layout = "LAYOUT_WIDE";
  presentation.author = "Office AI 助手";
  presentation.subject = title;
  presentation.title = title;
  presentation.company = "Office AI 助手";
  presentation.theme = {
    headFontFace: "Microsoft YaHei",
    bodyFontFace: "Microsoft YaHei",
  };
  const slide = presentation.addSlide();
  slide.background = { color: "FFFFFF" };
  slide.addText(title, {
    x: 0.8,
    y: 1.55,
    w: 11.7,
    h: 0.8,
    fontFace: "Microsoft YaHei",
    fontSize: 30,
    bold: true,
    color: "111827",
    align: "center",
    margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 1.2,
      y: 2.75,
      w: 10.9,
      h: 0.5,
      fontFace: "Microsoft YaHei",
      fontSize: 18,
      color: "4B5563",
      align: "center",
      margin: 0,
    });
  }
  await presentation.writeFile({ fileName: outputPath });
  return presentationDone(input, outputPath, ["ppt/presentation.xml", "ppt/slides/slide1.xml"], "已创建基础 PPTX 演示文稿");
}

async function applyTheme(input: PresentationAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const accentColor = normalizeColor(input.params?.accentColor);
  const changedParts: string[] = [];
  const slidePartNames = Object.keys(zip.files).filter((name) => PRESENTATION_SLIDE_RE.test(name)).sort();

  for (const partName of slidePartNames) {
    const part = zip.file(partName);
    if (!part) continue;
    const xml = await part.async("text");
    const nextXml = applyRunColor(xml, accentColor);
    if (nextXml !== xml) {
      zip.file(partName, nextXml);
      changedParts.push(partName);
    }
  }

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return presentationDone(input, outputPath, changedParts, "已应用 PPT 主题强调色");
}

async function deleteSlides(input: PresentationAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const presentationPart = zip.file(PRESENTATION_PART);
  const relsPart = zip.file(PRESENTATION_RELS_PART);
  if (!presentationPart) throw new Error(`找不到 PPT 主部件: ${PRESENTATION_PART}`);
  if (!relsPart) throw new Error(`找不到 PPT 关系部件: ${PRESENTATION_RELS_PART}`);

  const presentationXml = await presentationPart.async("text");
  const relsXml = await relsPart.async("text");
  const slideEntries = collectSlideEntries(presentationXml, relsXml);
  const indexes = resolveDeleteSlideIndexes(input, slideEntries.length);
  const deleteIndexSet = new Set(indexes);
  const deleteRelIds = new Set<string>();
  const deletedParts: string[] = [];

  for (const entry of slideEntries) {
    if (!deleteIndexSet.has(entry.index)) continue;
    deleteRelIds.add(entry.relId);
    if (entry.partName) {
      zip.remove(entry.partName);
      deletedParts.push(entry.partName);
    }
  }

  const nextPresentationXml = presentationXml.replace(/<p:sldId\b[^>]*(?:\/>|><\/p:sldId>)/g, (slideIdXml) => {
    const relId = getXmlAttr(slideIdXml, "r:id");
    return relId && deleteRelIds.has(relId) ? "" : slideIdXml;
  });
  const nextRelsXml = relsXml.replace(/<Relationship\b[^>]*(?:\/>|><\/Relationship>)/g, (relationshipXml) => {
    const relId = getXmlAttr(relationshipXml, "Id");
    return relId && deleteRelIds.has(relId) ? "" : relationshipXml;
  });
  const nextContentTypesXml = await removeSlideContentTypeOverrides(zip, deletedParts);

  zip.file(PRESENTATION_PART, nextPresentationXml);
  zip.file(PRESENTATION_RELS_PART, nextRelsXml);
  if (nextContentTypesXml) zip.file("[Content_Types].xml", nextContentTypesXml);

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return presentationDone(input, outputPath, deletedParts, `已删除 ${deletedParts.length} 张幻灯片`);
}

async function addSlides(input: PresentationAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const presentationPart = zip.file(PRESENTATION_PART);
  const relsPart = zip.file(PRESENTATION_RELS_PART);
  if (!presentationPart) throw new Error(`找不到 PPT 主部件: ${PRESENTATION_PART}`);
  if (!relsPart) throw new Error(`找不到 PPT 关系部件: ${PRESENTATION_RELS_PART}`);

  const slides = normalizeSlidesParam(input.params);
  if (slides.length === 0) {
    throw new Error("addSlide 需要 params.title/body 或 params.slides");
  }

  let presentationXml = await presentationPart.async("text");
  let relsXml = await relsPart.async("text");
  let contentTypesXml = await contentTypesXmlOrDefault(zip);
  let nextSlideNumber = nextSlidePartNumber(zip);
  let nextSlideId = nextPresentationSlideId(presentationXml);
  let nextRelNumber = nextRelationshipNumber(relsXml);
  const changedParts = [PRESENTATION_PART, PRESENTATION_RELS_PART, CONTENT_TYPES_PART];

  for (const slide of slides) {
    const slidePartName = `ppt/slides/slide${nextSlideNumber}.xml`;
    const relId = `rId${nextRelNumber}`;
    zip.file(slidePartName, contentSlideXml(slide.title, slide.body, slide.layout));
    zip.file(`ppt/slides/_rels/slide${nextSlideNumber}.xml.rels`, emptySlideRelsXml());
    presentationXml = insertSlideId(presentationXml, nextSlideId, relId);
    relsXml = insertPresentationRelationship(relsXml, relId, `slides/slide${nextSlideNumber}.xml`);
    contentTypesXml = ensureSlideContentType(contentTypesXml, slidePartName);
    changedParts.push(slidePartName);
    nextSlideNumber++;
    nextSlideId++;
    nextRelNumber++;
  }

  zip.file(PRESENTATION_PART, presentationXml);
  zip.file(PRESENTATION_RELS_PART, relsXml);
  zip.file(CONTENT_TYPES_PART, contentTypesXml);

  const outputPath = input.outputPath || input.filePath;
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return presentationDone(input, outputPath, changedParts, `已添加 ${slides.length} 张幻灯片`);
}

function applyRunColor(xml: string, color: string): string {
  return xml.replace(/<a:r\b[\s\S]*?<\/a:r>/g, (runXml) => {
    const fillXml = `<a:solidFill><a:srgbClr val="${color}" /></a:solidFill>`;
    if (/<a:rPr\b[\s\S]*?<\/a:rPr>/.test(runXml)) {
      return runXml.replace(/<a:rPr\b[^>]*>/, (tag) => `${tag}${fillXml}`);
    }
    return runXml.replace(/<a:r\b[^>]*>/, (tag) => `${tag}<a:rPr>${fillXml}</a:rPr>`);
  });
}

function resolveDeleteSlideIndexes(input: PresentationAdvancedActionInput, slideCount: number): number[] {
  const explicitSlides = arrayNumberParam(input.params?.slides);
  const range = explicitSlides.length > 0
    ? explicitSlides
    : rangeFromParams(input.params) || rangeFromTarget(input.target);
  if (!range || range.length === 0) {
    throw new Error("deleteSlides 需要 params.slides、params.from/to 或 target: slide:2-6");
  }

  const indexes = [...new Set(range.map((value) => Math.floor(value)).filter((value) => value >= 1))].sort((a, b) => a - b);
  if (indexes.length === 0) throw new Error("deleteSlides 未解析到有效幻灯片序号");
  const overflow = indexes.find((index) => index > slideCount);
  if (overflow) throw new Error(`幻灯片序号超出范围: ${overflow}`);
  if (indexes.length >= slideCount) throw new Error("deleteSlides 至少需要保留一张幻灯片");
  return indexes;
}

function rangeFromParams(params?: Record<string, unknown>): number[] | undefined {
  const from = numberParam(params, "from") ?? numberParam(params, "start");
  const to = numberParam(params, "to") ?? numberParam(params, "end") ?? from;
  return from && to ? buildNumberRange(from, to) : undefined;
}

function rangeFromTarget(target?: string): number[] | undefined {
  if (!target) return undefined;
  const match = target.match(/^slides?:\s*(\d+)(?:\s*-\s*(\d+))?$/i);
  if (!match) return undefined;
  const from = Number(match[1]);
  const to = Number(match[2] || match[1]);
  return buildNumberRange(from, to);
}

function buildNumberRange(from: number, to: number): number[] {
  const start = Math.floor(Math.min(from, to));
  const end = Math.floor(Math.max(from, to));
  return Array.from({ length: end - start + 1 }, (_, index) => start + index);
}

function arrayNumberParam(value: unknown): number[] {
  return Array.isArray(value)
    ? value.map((item) => Number(item)).filter((item) => Number.isFinite(item))
    : [];
}

function numberParam(params: Record<string, unknown> | undefined, key: string): number | undefined {
  const value = params?.[key];
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

async function removeSlideContentTypeOverrides(zip: JSZip, deletedParts: string[]): Promise<string | undefined> {
  const contentTypesPart = zip.file(CONTENT_TYPES_PART);
  if (!contentTypesPart || deletedParts.length === 0) return undefined;
  const deletedNames = new Set(deletedParts.map((partName) => `/${partName}`));
  const xml = await contentTypesPart.async("text");
  return xml.replace(/<Override\b[^>]*(?:\/>|><\/Override>)/g, (overrideXml) => {
    const partName = getXmlAttr(overrideXml, "PartName");
    return partName && deletedNames.has(partName) ? "" : overrideXml;
  });
}

function normalizeColor(value: unknown): string {
  return typeof value === "string" && /^[0-9a-fA-F]{6}$/.test(value) ? value.toUpperCase() : "1F4E79";
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  return typeof params?.[key] === "string" ? params[key] : undefined;
}

function defaultOutputPath(filePath: string): string {
  const ext = path.extname(filePath);
  return path.join(path.dirname(filePath), `${path.basename(filePath, ext)}-advanced${ext}`);
}
