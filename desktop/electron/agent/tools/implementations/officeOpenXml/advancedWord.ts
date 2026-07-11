/**
 * Word Open XML 高级操作
 *
 * 关联模块：
 * - officeCore/officeActionAdapter.ts: 将统一 Office action 路由到本模块。
 * - tableStyler.ts: 复用 Word 表格样式处理。
 */

import { readFile, writeFile } from "fs/promises";
import path from "path";
import JSZip from "jszip";
import { failedResult, needsComResult, unsupportedResult } from "../../officeCore/results";
import type {
  OfficeActionKind,
  OfficeActionResult,
  OfficeActionValidation,
} from "../../officeCore/types";
import {
  decodeXmlText as unescapeXml,
  escapeXmlText as escapeXml,
  parseXmlAttributes,
} from "../../../shared/xmlEntities";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";
import { createOpenXmlDoneResult } from "./actionResult";
import type { OfficeOpenXmlTableStylePreset } from "./types";

export interface WordAdvancedActionInput {
  operation: string;
  filePath: string;
  outputPath?: string;
  target?: string;
  action?: OfficeActionKind;
  params?: Record<string, unknown>;
}

const WORD_DOCUMENT_PART = "word/document.xml";
const WORD_DOCUMENT_RELS_PART = "word/_rels/document.xml.rels";
const CONTENT_TYPES_PART = "[Content_Types].xml";
const WORDPROCESSINGML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const OFFICE_RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TABLE_STYLES = new Set<OfficeOpenXmlTableStylePreset>(["professional", "compact", "financial"]);
const wordDone = createOpenXmlDoneResult("word");

export async function applyWordAdvancedAction(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  try {
    if (input.operation === "createDocument") {
      return await createDocument(input);
    }
    if (input.operation === "applyHeadingStyles") {
      return await applyHeadingStyles(input);
    }
    if (input.operation === "styleTables") {
      return await styleTables(input);
    }
    if (input.operation === "setHeaderFooter") {
      return await setHeaderFooter(input);
    }
    if (input.operation === "insertOrUpdateToc" || input.operation === "insertOrReplaceImage") {
      return needsComResult({
        app: "word",
        action: input.action || "insert",
        operation: input.operation,
        filePath: input.filePath,
        outputPath: input.outputPath,
        target: input.target,
        summary: `${input.operation} 需要 Word 刷新字段或维护关系图，需显式 COM 兜底`,
      });
    }

    return unsupportedResult({
      app: "word",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
      summary: `暂不支持 Word Open XML 高级操作: ${input.operation}`,
    });
  } catch (error) {
    return failedResult({
      app: "word",
      action: input.action || "edit",
      operation: input.operation,
      filePath: input.filePath,
      outputPath: input.outputPath,
      target: input.target,
    }, error);
  }
}

async function createDocument(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const outputPath = input.outputPath || input.filePath;
  const title = typeof input.params?.title === "string" ? input.params.title : "";
  const body = normalizeParagraphs(input.params?.paragraphs ?? input.params?.text ?? input.params?.body);
  const zip = new JSZip();

  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`);
  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
  zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file("word/styles.xml", basicStylesXml());
  zip.file(WORD_DOCUMENT_PART, documentXml(title, body));

  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return wordDone(input, outputPath, [WORD_DOCUMENT_PART, "word/styles.xml"], "已使用内置 Open XML 创建 Word 文档", {
    engine: "openxml",
    operation: "createDocument",
    filePath: input.filePath,
    outputPath,
    title,
    paragraphCount: body.length,
  });
}

async function applyHeadingStyles(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const part = zip.file(WORD_DOCUMENT_PART);
  if (!part) throw new Error(`找不到 Word 正文部件: ${WORD_DOCUMENT_PART}`);

  const startsWith = typeof input.params?.startsWith === "string" ? input.params.startsWith : "";
  const level = normalizeHeadingLevel(input.params?.level);
  const xml = await part.async("text");
  let changed = false;
  const nextXml = xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    const text = paragraphText(paragraphXml);
    if (startsWith && !text.startsWith(startsWith)) return paragraphXml;
    changed = true;
    return applyParagraphStyle(paragraphXml, `Heading${level}`);
  });
  zip.file(WORD_DOCUMENT_PART, nextXml);

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  return wordDone(input, outputPath, changed ? [WORD_DOCUMENT_PART] : [], "已应用 Word 标题样式");
}

async function styleTables(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const style = normalizeTableStyle(input.params?.style);
  const result = await applyOfficeOpenXmlTableStyle({
    filePath: input.filePath,
    outputPath: input.outputPath,
    target: input.target,
    style,
  });
  return wordDone(input, result.outputPath, result.changedParts, "已应用 Word 表格样式", result);
}

async function setHeaderFooter(input: WordAdvancedActionInput): Promise<OfficeActionResult> {
  const zip = await JSZip.loadAsync(await readFile(input.filePath));
  const kind = input.operation === "setHeaderFooter" && input.params?.kind === "footer" ? "footer" : "header";
  const text = typeof input.params?.text === "string" ? input.params.text : "";
  const documentPart = zip.file(WORD_DOCUMENT_PART);
  const contentTypesPart = zip.file(CONTENT_TYPES_PART);
  if (!documentPart) throw new Error(`找不到 Word 正文部件: ${WORD_DOCUMENT_PART}`);
  if (!contentTypesPart) throw new Error(`找不到 Word 内容类型部件: ${CONTENT_TYPES_PART}`);

  let documentXml = await documentPart.async("text");
  let contentTypesXml = await contentTypesPart.async("text");
  let relsXml = zip.file(WORD_DOCUMENT_RELS_PART)
    ? await zip.file(WORD_DOCUMENT_RELS_PART)!.async("text")
    : relationshipsXml();

  const existingReferenceId = findDefaultReferenceId(documentXml, kind);
  const existingRelationship = existingReferenceId
    ? findRelationship(relsXml, existingReferenceId)
    : undefined;
  const canReuseRelationship = existingRelationship
    && existingRelationship.type.endsWith(`/relationships/${kind}`)
    && existingRelationship.target
    && existingRelationship.targetMode?.toLowerCase() !== "external";

  const relationshipId = canReuseRelationship
    ? existingRelationship.id
    : nextRelationshipId(relsXml);
  const relationshipTarget = canReuseRelationship
    ? existingRelationship.target
    : nextHeaderFooterTarget(zip, kind);
  const partName = normalizeWordRelationshipTarget(relationshipTarget);

  if (!canReuseRelationship) {
    relsXml = appendRelationship(relsXml, {
      id: relationshipId,
      type: `${OFFICE_RELATIONSHIPS_NS}/${kind}`,
      target: relationshipTarget,
    });
  }

  documentXml = ensureRelationshipNamespace(documentXml);
  documentXml = upsertDefaultSectionReference(documentXml, kind, relationshipId);
  contentTypesXml = upsertContentTypeOverride(contentTypesXml, partName, kind);

  zip.file(partName, wordTextPartXml(kind, text));
  zip.file(WORD_DOCUMENT_PART, documentXml);
  zip.file(WORD_DOCUMENT_RELS_PART, relsXml);
  zip.file(CONTENT_TYPES_PART, contentTypesXml);

  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  await writeFile(outputPath, await zip.generateAsync({ type: "nodebuffer" }));
  const validation = await validateHeaderFooterOutput({
    outputPath,
    kind,
    partName,
    relationshipId,
    text,
  });
  const result = wordDone(
    input,
    outputPath,
    [partName, WORD_DOCUMENT_PART, WORD_DOCUMENT_RELS_PART, CONTENT_TYPES_PART],
    kind === "footer" ? "已写入 Word 页脚" : "已写入 Word 页眉",
    undefined,
    validation
  );
  if (validation.ok) return result;
  return {
    ...result,
    status: "failed",
    error: validation.checks
      .filter((check) => !check.ok)
      .map((check) => check.message)
      .join("；"),
  };
}

function paragraphText(paragraphXml: string): string {
  return [...paragraphXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function applyParagraphStyle(paragraphXml: string, styleId: string): string {
  const styleXml = `<w:pStyle w:val="${styleId}" />`;
  if (/<w:pPr\b[\s\S]*?<\/w:pPr>/.test(paragraphXml)) {
    return paragraphXml.replace(/<w:pPr\b[\s\S]*?<\/w:pPr>/, (pPrXml) => {
      if (/<w:pStyle\b[^>]*(?:\/>|><\/w:pStyle>)/.test(pPrXml)) {
        return pPrXml.replace(/<w:pStyle\b[^>]*(?:\/>|><\/w:pStyle>)/, styleXml);
      }
      return pPrXml.replace(/<w:pPr\b[^>]*>/, (tag) => `${tag}${styleXml}`);
    });
  }
  return paragraphXml.replace(/<w:p\b[^>]*>/, (tag) => `${tag}<w:pPr>${styleXml}</w:pPr>`);
}

function wordTextPartXml(kind: "header" | "footer", text: string): string {
  const root = kind === "footer" ? "w:ftr" : "w:hdr";
  return `<${root} xmlns:w="${WORDPROCESSINGML_NS}"><w:p><w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p></${root}>`;
}

function relationshipsXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
}

function findDefaultReferenceId(
  documentXml: string,
  kind: "header" | "footer"
): string | undefined {
  const referenceRe = new RegExp(`<w:${kind}Reference\\b[^>]*>`, "g");
  let match: RegExpExecArray | null;
  while ((match = referenceRe.exec(documentXml))) {
    const attrs = parseXmlAttributes(match[0]);
    if ((attrs["w:type"] || attrs.type) === "default") {
      return attrs["r:id"] || attrs.id;
    }
  }
  return undefined;
}

function findRelationship(
  relsXml: string,
  relationshipId: string
): { id: string; type: string; target: string; targetMode?: string } | undefined {
  const relationshipRe = /<Relationship\b[^>]*>/g;
  let match: RegExpExecArray | null;
  while ((match = relationshipRe.exec(relsXml))) {
    const attrs = parseXmlAttributes(match[0]);
    if (attrs.Id === relationshipId) {
      return {
        id: attrs.Id,
        type: attrs.Type || "",
        target: attrs.Target || "",
        targetMode: attrs.TargetMode,
      };
    }
  }
  return undefined;
}

function nextRelationshipId(relsXml: string): string {
  let maxId = 0;
  for (const match of relsXml.matchAll(/\bId=["']rId(\d+)["']/g)) {
    maxId = Math.max(maxId, Number(match[1]));
  }
  return `rId${maxId + 1}`;
}

function nextHeaderFooterTarget(zip: JSZip, kind: "header" | "footer"): string {
  let maxIndex = 0;
  const partRe = new RegExp(`^word/${kind}(\\d+)\\.xml$`, "i");
  for (const name of Object.keys(zip.files)) {
    const match = partRe.exec(name);
    if (match) maxIndex = Math.max(maxIndex, Number(match[1]));
  }
  return `${kind}${maxIndex + 1}.xml`;
}

function appendRelationship(
  relsXml: string,
  relationship: { id: string; type: string; target: string }
): string {
  const addition = `<Relationship Id="${relationship.id}" Type="${relationship.type}" Target="${relationship.target}"/>`;
  if (relsXml.includes("</Relationships>")) {
    return relsXml.replace("</Relationships>", `${addition}</Relationships>`);
  }
  const selfClosing = /<Relationships\b[^>]*\/>/.exec(relsXml);
  if (selfClosing) {
    const openingTag = selfClosing[0].replace(/\/>$/, ">");
    return relsXml.replace(selfClosing[0], `${openingTag}${addition}</Relationships>`);
  }
  return `${relationshipsXml().replace("</Relationships>", `${addition}</Relationships>`)}`;
}

function normalizeWordRelationshipTarget(target: string): string {
  const normalized = target.replace(/\\/g, "/").replace(/^\/+/, "");
  if (normalized.startsWith("word/")) return normalized;
  return path.posix.join("word", normalized);
}

function ensureRelationshipNamespace(documentXml: string): string {
  const documentTag = /<w:document\b[^>]*>/.exec(documentXml)?.[0];
  if (!documentTag || /\bxmlns:r\s*=/.test(documentTag)) return documentXml;
  return documentXml.replace(
    documentTag,
    documentTag.replace(/>$/, ` xmlns:r="${OFFICE_RELATIONSHIPS_NS}">`)
  );
}

function upsertDefaultSectionReference(
  documentXml: string,
  kind: "header" | "footer",
  relationshipId: string
): string {
  const referenceXml = `<w:${kind}Reference w:type="default" r:id="${relationshipId}"/>`;
  const referenceRe = new RegExp(
    `<w:${kind}Reference\\b[^>]*(?:\\/\\s*>|>[\\s\\S]*?<\\/w:${kind}Reference\\s*>)`,
    "g"
  );
  let updatedAnySection = false;

  let nextXml = documentXml.replace(
    /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g,
    (sectionXml) => {
      updatedAnySection = true;
      const withoutDefault = sectionXml.replace(referenceRe, (reference) => {
        const attrs = parseXmlAttributes(reference);
        return (attrs["w:type"] || attrs.type) === "default" ? "" : reference;
      });
      const openingTag = /^<w:sectPr\b[^>]*>/.exec(withoutDefault)?.[0];
      if (!openingTag) return withoutDefault;
      const body = withoutDefault.slice(
        openingTag.length,
        withoutDefault.length - "</w:sectPr>".length
      );
      return `${openingTag}${insertOrderedSectionReference(body, kind, referenceXml)}</w:sectPr>`;
    }
  );

  nextXml = nextXml.replace(/<w:sectPr\b[^>]*\/>/g, (sectionXml) => {
    updatedAnySection = true;
    const openingTag = sectionXml.replace(/\/>$/, ">");
    return `${openingTag}${referenceXml}</w:sectPr>`;
  });

  if (updatedAnySection) return nextXml;
  if (nextXml.includes("</w:body>")) {
    return nextXml.replace("</w:body>", `<w:sectPr>${referenceXml}</w:sectPr></w:body>`);
  }
  return nextXml;
}

function insertOrderedSectionReference(
  sectionBody: string,
  kind: "header" | "footer",
  referenceXml: string
): string {
  if (kind === "header") return `${referenceXml}${sectionBody}`;

  const leadingReferenceRe =
    /\s*<w:(?:header|footer)Reference\b[^>]*(?:\/\s*>|>[\s\S]*?<\/w:(?:header|footer)Reference\s*>)/y;
  let insertAt = 0;
  while (insertAt < sectionBody.length) {
    leadingReferenceRe.lastIndex = insertAt;
    const match = leadingReferenceRe.exec(sectionBody);
    if (!match) break;
    insertAt = leadingReferenceRe.lastIndex;
  }
  return `${sectionBody.slice(0, insertAt)}${referenceXml}${sectionBody.slice(insertAt)}`;
}

function upsertContentTypeOverride(
  contentTypesXml: string,
  partName: string,
  kind: "header" | "footer"
): string {
  const packagePartName = `/${partName}`;
  const contentType = `application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml`;
  const overrideRe = /<Override\b[^>]*\/?>/g;
  let found = false;
  const updated = contentTypesXml.replace(overrideRe, (override) => {
    const attrs = parseXmlAttributes(override);
    if (attrs.PartName !== packagePartName) return override;
    found = true;
    return setXmlAttribute(override, "ContentType", contentType);
  });
  if (found) return updated;

  const addition = `<Override PartName="${packagePartName}" ContentType="${contentType}"/>`;
  if (updated.includes("</Types>")) {
    return updated.replace("</Types>", `${addition}</Types>`);
  }
  const selfClosing = /<Types\b[^>]*\/>/.exec(updated);
  if (selfClosing) {
    const openingTag = selfClosing[0].replace(/\/>$/, ">");
    return updated.replace(selfClosing[0], `${openingTag}${addition}</Types>`);
  }
  return updated;
}

function setXmlAttribute(tagXml: string, name: string, value: string): string {
  const attributeRe = new RegExp(`\\b${name}\\s*=\\s*["'][^"']*["']`);
  if (attributeRe.test(tagXml)) {
    return tagXml.replace(attributeRe, `${name}="${value}"`);
  }
  return tagXml.replace(/\/?>$/, (end) => ` ${name}="${value}"${end}`);
}

async function validateHeaderFooterOutput(input: {
  outputPath: string;
  kind: "header" | "footer";
  partName: string;
  relationshipId: string;
  text: string;
}): Promise<OfficeActionValidation> {
  const zip = await JSZip.loadAsync(await readFile(input.outputPath));
  const partXml = await zip.file(input.partName)?.async("text");
  const documentXml = await zip.file(WORD_DOCUMENT_PART)?.async("text");
  const relsXml = await zip.file(WORD_DOCUMENT_RELS_PART)?.async("text");
  const contentTypesXml = await zip.file(CONTENT_TYPES_PART)?.async("text");
  const relationship = relsXml
    ? findRelationship(relsXml, input.relationshipId)
    : undefined;
  const referenceName = input.kind === "footer" ? "footerReference" : "headerReference";
  const contentType = `wordprocessingml.${input.kind}+xml`;

  const checks = [
    {
      name: "header-footer-part",
      ok: Boolean(
        partXml
        && partXml.includes(`xmlns:w="${WORDPROCESSINGML_NS}"`)
        && partXml.includes(escapeXml(input.text))
      ),
      message: `${input.partName} 已生成并包含目标文本`,
    },
    {
      name: "header-footer-relationship",
      ok: Boolean(
        relationship
        && relationship.type.endsWith(`/relationships/${input.kind}`)
        && normalizeWordRelationshipTarget(relationship.target) === input.partName
      ),
      message: `${WORD_DOCUMENT_RELS_PART} 已连接 ${input.partName}`,
    },
    {
      name: "section-reference",
      ok: Boolean(
        documentXml
        && documentXml.includes(
          `w:${referenceName} w:type="default" r:id="${input.relationshipId}"`
        )
      ),
      message: `${WORD_DOCUMENT_PART} 已引用 ${input.relationshipId}`,
    },
    {
      name: "content-type",
      ok: Boolean(
        contentTypesXml
        && contentTypesXml.includes(`PartName="/${input.partName}"`)
        && contentTypesXml.includes(contentType)
      ),
      message: `${CONTENT_TYPES_PART} 已声明 ${input.kind} 内容类型`,
    },
  ];

  return {
    ok: checks.every((check) => check.ok),
    checks,
  };
}

function normalizeParagraphs(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value.map((item) => String(item)).filter((item) => item.length > 0);
  }
  if (typeof value === "string" && value.length > 0) {
    return value.split(/\r?\n/).filter((item) => item.length > 0);
  }
  return [];
}

function documentXml(title: string, paragraphs: string[]): string {
  const parts = [
    title ? paragraphXml(title, "Title") : "",
    ...paragraphs.map((text) => paragraphXml(text)),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${parts}<w:sectPr /></w:body>
</w:document>`;
}

function paragraphXml(text: string, styleId?: string): string {
  const style = styleId ? `<w:pPr><w:pStyle w:val="${styleId}" /></w:pPr>` : "";
  return `<w:p>${style}<w:r><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function basicStylesXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>`;
}

function normalizeHeadingLevel(value: unknown): number {
  return typeof value === "number" && value >= 1 && value <= 9 ? Math.floor(value) : 1;
}

function normalizeTableStyle(value: unknown): OfficeOpenXmlTableStylePreset {
  return typeof value === "string" && TABLE_STYLES.has(value as OfficeOpenXmlTableStylePreset)
    ? value as OfficeOpenXmlTableStylePreset
    : "professional";
}

function defaultOutputPath(filePath: string): string {
  const ext = path.extname(filePath);
  return path.join(path.dirname(filePath), `${path.basename(filePath, ext)}-advanced${ext}`);
}
