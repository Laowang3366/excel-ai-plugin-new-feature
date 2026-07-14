import { trustedIpcMain as ipcMain } from "../shared/trustedIpc";
import { createAIClient } from "../agent/providers/aiClient";
import { parseFilesLocally } from "../agent/tools/executors/localDocumentParser";
import { validateInput, OcrRecognizeInput } from "../shared/ipcSchemas";
import { assertAuthorizedPath, createPathAuthorizer } from "./ipcPathSecurity";
import { getActiveAIConfig, getRuntimeSettingValue } from "./settingsManager";
import { parseFilesWithMineru, parseFilesWithMineruAgent, type MineruParsedDocument } from "./mineruOcr";
import {
  buildInvoiceFieldFallback,
  buildRowsFromFields,
  isLikelyInvoiceText,
  mergeInvoiceFields,
} from "./invoiceFieldExtraction";
import {
  normalizeOcrMode,
  isLikelyInvoiceFileList,
  normalizePlainOcrText,
  type OcrInvoiceItem,
  type OcrVisionResult,
} from "./ocrModeDetection";

export function registerOcrIpcHandler(pathAuthorizer: ReturnType<typeof createPathAuthorizer>): void {
  try {
    ipcMain.removeHandler("ocr:recognize");
  } catch {
    // Handler may not exist on first registration.
  }

  ipcMain.handle("ocr:recognize", async (_event, mode: unknown, filePaths: unknown) => {
    try {
      const validated = validateInput(OcrRecognizeInput, { mode, filePaths });
      const authorizedFilePaths = validated.filePaths.map((filePath) =>
        assertAuthorizedPath(pathAuthorizer, filePath)
      );
      return await recognizeWithOcrFallbacks(validated.mode, authorizedFilePaths);
    } catch (err: any) {
      return emptyOcrResult(normalizeOcrMode(mode), [err?.message || "OCR 识别失败"]);
    }
  });
}

async function recognizeWithOcrFallbacks(
  rawMode: unknown,
  rawFilePaths: unknown,
): Promise<OcrVisionResult> {
  const mode = normalizeOcrMode(rawMode);
  const filePaths = normalizeOcrFilePaths(rawFilePaths);
  const effectiveMode = mode === "invoice" || isLikelyInvoiceFileList(filePaths) ? "invoice" : "image";

  const parsed = await parseFilesWithOcrFallbacks(filePaths);
  if (!hasAnyUsefulParsedDocument(parsed.documents)) {
    return emptyOcrResult(effectiveMode, [
      "未提取到可用 OCR 文本或表格，无法抽取字段",
      ...parsed.errors,
      ...formatParsedDocumentErrors(parsed.documents),
    ]);
  }

  const result = effectiveMode === "invoice" || isLikelyInvoiceDocuments(parsed.documents)
    ? await extractInvoiceFieldsFromMineruDocuments(parsed.documents)
    : buildMineruOcrResult(parsed.documents);

  return {
    ...result,
    errors: [
      ...result.errors,
      ...formatParsedDocumentErrors(parsed.documents),
    ],
  };
}

async function parseFilesWithOcrFallbacks(
  filePaths: string[],
): Promise<{ documents: MineruParsedDocument[]; errors: string[] }> {
  const errors: string[] = [];
  const mineruToken = getConfiguredMineruToken();
  const selected: Array<MineruParsedDocument | undefined> = new Array(filePaths.length);
  let unresolved = filePaths.map((_, index) => index);

  if (mineruToken) {
    try {
      const documents = await parseFilesWithMineru(filePaths, mineruToken);
      unresolved = mergeUsefulDocuments(selected, unresolved, documents);
      if (unresolved.length > 0) errors.push(formatMineruDocumentErrors(documents) || "MinerU 标准解析存在未完成文件");
    } catch (error: any) {
      errors.push(`MinerU 标准解析失败：${error?.message || "未知错误"}`);
    }
  }

  if (unresolved.length > 0) {
    const pendingIndices = unresolved;
    try {
      const documents = await parseFilesWithMineruAgent(pendingIndices.map((index) => filePaths[index]));
      unresolved = mergeUsefulDocuments(selected, pendingIndices, documents);
      if (unresolved.length > 0) errors.push(formatMineruDocumentErrors(documents) || "MinerU 免费解析存在未完成文件");
    } catch (error: any) {
      errors.push(`MinerU 免费解析失败：${error?.message || "未知错误"}`);
    }
  }

  if (unresolved.length > 0) {
    const pendingIndices = unresolved;
    const localDocuments = await parseFilesLocally(pendingIndices.map((index) => filePaths[index]));
    for (let index = 0; index < pendingIndices.length; index++) {
      selected[pendingIndices[index]] = localDocuments[index];
    }
  }

  const documents = selected.filter((document): document is MineruParsedDocument => Boolean(document));
  return {
    documents,
    errors: hasAnyUsefulParsedDocument(documents) ? [] : errors,
  };
}

function mergeUsefulDocuments(
  selected: Array<MineruParsedDocument | undefined>,
  targetIndices: number[],
  documents: MineruParsedDocument[],
): number[] {
  const unresolved: number[] = [];
  for (let index = 0; index < targetIndices.length; index++) {
    const document = documents[index];
    if (document && hasAnyUsefulParsedDocument([document])) {
      selected[targetIndices[index]] = document;
    } else {
      unresolved.push(targetIndices[index]);
    }
  }
  return unresolved;
}

function hasAnyUsefulParsedDocument(documents: Array<{ text: string; rows: string[][] }>): boolean {
  return documents.some((document) => document.text.trim().length > 0 || document.rows.length > 0);
}

function formatParsedDocumentErrors(documents: Array<{ filename: string; error?: string }>): string[] {
  return documents
    .filter((document) => document.error && !/^local_ocr_unsupported|local_unsupported|local_empty$/.test(document.error))
    .map((document) => `${document.filename}: ${document.error}`);
}

function normalizeOcrFilePaths(rawFilePaths: unknown): string[] {
  if (!Array.isArray(rawFilePaths)) {
    throw new Error("OCR 文件列表必须是数组");
  }

  const filePaths = rawFilePaths
    .filter((value): value is string => typeof value === "string")
    .map((value) => value.trim())
    .filter(Boolean);
  if (filePaths.length === 0) {
    throw new Error("请先选择要识别的图片或 PDF");
  }

  return filePaths;
}

function isLikelyInvoiceDocuments(documents: MineruParsedDocument[]): boolean {
  return documents.some((document) =>
    isLikelyInvoiceText(`${document.filename}\n${document.text}`)
  );
}

function getConfiguredMineruToken(): string {
  const configured =
    getRuntimeSettingValue("mineruApiToken") ||
    getRuntimeSettingValue("ocrMineruApiToken");
  const tokenFromSettings = typeof configured === "string" ? configured.trim() : "";
  return tokenFromSettings || (process.env.MINERU_API_TOKEN || "").trim();
}

function buildMineruOcrResult(documents: MineruParsedDocument[]): OcrVisionResult {
  const text = buildCombinedMineruText(documents);
  return {
    kind: "image",
    text,
    rows: documents.flatMap((document) => document.rows),
    fields: {},
    invoices: [],
    errors: normalizeStringArray(formatMineruDocumentErrors(documents).split("\n")),
  };
}

async function extractInvoiceFieldsFromMineruDocuments(
  documents: MineruParsedDocument[],
): Promise<OcrVisionResult> {
  const fallback = buildMineruInvoiceFallbackResult(documents);
  try {
    const aiClient = createAIClient(getActiveAIConfig());
    const result = await aiClient.chat({
      messages: [{
        role: "user",
        content: buildMineruInvoicePrompt(documents),
      }],
      maxTokens: 4000,
      temperature: 0,
      reasoningMode: "off",
    });
    const normalized = normalizeOcrVisionResult("invoice", result.content || "");
    const extractedInvoices = normalizeExtractedInvoiceItems(normalized, documents);
    const mergedInvoices = mergeMineruInvoiceItems(extractedInvoices, documents);
    return {
      kind: "invoice",
      text: normalized.text || fallback.text,
      rows: normalized.rows.length > 0 ? normalized.rows : fallback.rows,
      fields: Object.keys(normalized.fields).length > 0 ? normalized.fields : fallback.fields,
      invoices: mergedInvoices.length > 0 ? mergedInvoices : fallback.invoices,
      errors: [
        ...fallback.errors,
        ...normalized.errors,
      ],
    };
  } catch (error: any) {
    return {
      ...fallback,
      errors: [
        ...fallback.errors,
        `发票字段抽取失败，已保留 MinerU OCR 文本：${error?.message || "未知错误"}`,
      ],
    };
  }
}

function normalizeExtractedInvoiceItems(
  normalized: OcrVisionResult,
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  if (normalized.invoices.length > 0) return normalized.invoices;
  if (Object.keys(normalized.fields).length === 0) return [];
  if (documents.length <= 1) {
    const document = documents[0];
    return [{
      filename: document?.filename || "识别文本",
      text: normalized.text || document?.text || "",
      fields: normalized.fields,
      rows: normalized.rows,
    }];
  }
  return [{
    filename: documents[0]?.filename || "识别文本",
    text: normalized.text || documents[0]?.text || "",
    fields: normalized.fields,
    rows: normalized.rows,
  }];
}

function buildMineruInvoiceFallbackResult(documents: MineruParsedDocument[]): OcrVisionResult {
  const extracted = buildInvoiceFieldFallback(documents);
  return {
    kind: "invoice",
    text: buildCombinedMineruText(documents),
    rows: extracted.rows.length > 0 ? extracted.rows : documents.flatMap((document) => document.rows),
    fields: extracted.fields,
    invoices: extracted.invoices,
    errors: normalizeStringArray(formatMineruDocumentErrors(documents).split("\n")),
  };
}

function buildMineruInvoicePrompt(documents: MineruParsedDocument[]): string {
  return [
    "下面是 MinerU 通用 OCR/版面解析得到的发票 Markdown 文本。",
    "请只基于这些文本抽取发票字段，不要编造缺失信息。",
    "只返回严格 JSON，不要 Markdown，不要解释。",
    "JSON 结构必须是：",
    "{\"kind\":\"invoice\",\"text\":\"合并后的可读文本\",\"rows\":[[\"列1\",\"列2\"]],\"fields\":{\"字段\":\"值\"},\"invoices\":[{\"filename\":\"文件名\",\"text\":\"文本\",\"fields\":{\"发票号码\":\"\",\"开票日期\":\"\",\"购买方名称\":\"\",\"销售方名称\":\"\",\"金额\":\"\",\"税额\":\"\",\"价税合计\":\"\"},\"rows\":[[\"列1\",\"列2\"]]}],\"errors\":[]}",
    "字段优先包含：发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、发票类型、校验码、备注。",
    "每个输入文件都要在 invoices 中返回一项；未识别字段填空字符串。",
    "",
    buildLimitedMineruSource(documents),
  ].join("\n");
}

function buildCombinedMineruText(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.text.trim())
    .map((document) => `## ${document.filename}\n${document.text.trim()}`)
    .join("\n\n");
}

function buildLimitedMineruSource(documents: MineruParsedDocument[]): string {
  const perDocumentLimit = 12_000;
  const totalLimit = 40_000;
  let used = 0;
  const sections: string[] = [];

  for (const document of documents) {
    if (!document.text.trim()) continue;
    const remaining = totalLimit - used;
    if (remaining <= 0) break;
    const clippedText = document.text.trim().slice(0, Math.min(perDocumentLimit, remaining));
    used += clippedText.length;
    sections.push(`### 文件：${document.filename}\n${clippedText}`);
  }

  return sections.join("\n\n");
}

function mergeMineruInvoiceItems(
  extractedInvoices: OcrInvoiceItem[],
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  const fallbackInvoices = buildInvoiceFieldFallback(documents).invoices;
  if (documents.length === 0) return extractedInvoices;
  return documents.map((document, index) => {
    const extracted = extractedInvoices.find((invoice) => invoice.filename === document.filename)
      || extractedInvoices[index];
    const fallback = fallbackInvoices.find((invoice) => invoice.filename === document.filename)
      || fallbackInvoices[index];
    const fields = mergeInvoiceFields(fallback?.fields, extracted?.fields);
    return {
      filename: extracted?.filename || document.filename,
      text: extracted?.text || document.text,
      fields,
      rows: extracted?.rows?.length
        ? extracted.rows
        : Object.keys(fields).length > 0
        ? buildRowsFromFields([fields])
        : fallback?.rows?.length
        ? fallback.rows
        : document.rows,
      error: document.error || extracted?.error,
    };
  });
}

function formatMineruDocumentErrors(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.error)
    .map((document) => `${document.filename}: ${document.error}`)
    .join("\n");
}

export function normalizeOcrVisionResult(mode: "image" | "invoice", content: string): OcrVisionResult {
  const parsed = parseJsonObject(content);
  if (!parsed) {
    return normalizePlainOcrText(mode, content);
  }

  const result = parsed as Record<string, unknown>;
  const invoices = normalizeInvoices(result.invoices);
  return {
    kind: mode,
    text: typeof result.text === "string" ? result.text : invoices.map((item) => item.text).filter(Boolean).join("\n"),
    rows: normalizeRows(result.rows),
    fields: normalizeFields(result.fields),
    invoices,
    errors: normalizeStringArray(result.errors),
  };
}

function parseJsonObject(content: string): Record<string, unknown> | null {
  const trimmed = content.trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  const candidate = fenced ? fenced[1].trim() : trimmed;
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  try {
    const parsed = JSON.parse(candidate.slice(start, end + 1));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? parsed as Record<string, unknown>
      : null;
  } catch {
    return null;
  }
}

function emptyOcrResult(mode: "image" | "invoice", errors: string[] = []): OcrVisionResult {
  return {
    kind: mode,
    text: "",
    rows: [],
    fields: {},
    invoices: [],
    errors,
  };
}

function normalizeInvoices(value: unknown): OcrInvoiceItem[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((item): OcrInvoiceItem[] => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return [];
    const raw = item as Record<string, unknown>;
    return [{
      filename: typeof raw.filename === "string" ? raw.filename : "",
      text: typeof raw.text === "string" ? raw.text : "",
      fields: normalizeFields(raw.fields),
      rows: normalizeRows(raw.rows),
      error: typeof raw.error === "string" ? raw.error : undefined,
    }];
  });
}

function normalizeFields(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const fields: Record<string, string> = {};
  for (const [key, rawValue] of Object.entries(value)) {
    if (!key.trim()) continue;
    fields[key] = rawValue === null || rawValue === undefined
      ? ""
      : typeof rawValue === "string"
      ? rawValue
      : String(rawValue);
  }
  return fields;
}

function normalizeRows(value: unknown): string[][] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((row): string[][] => {
    if (!Array.isArray(row)) return [];
    return [row.map((cell) => cell === null || cell === undefined ? "" : String(cell))];
  });
}

function normalizeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => item.trim())
    .filter(Boolean);
}
