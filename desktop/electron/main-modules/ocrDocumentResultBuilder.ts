import { createAIClient } from "../agent/providers/aiClient";
import { assertRemoteDataProcessingAllowed } from "../shared/egressPolicy";
import {
  buildInvoiceFieldFallback,
  buildRowsFromFields,
  isLikelyInvoiceText,
  mergeInvoiceFields,
} from "./invoiceFieldExtraction";
import type { MineruParsedDocument } from "./mineruOcr";
import {
  normalizePlainOcrText,
  type OcrInvoiceItem,
  type OcrVisionResult,
} from "./ocrModeDetection";
import { getActiveAIConfig } from "./settingsManager";

export async function buildOcrResultFromDocuments(
  documents: MineruParsedDocument[],
  requestedMode: "image" | "invoice",
  remoteEnabled: boolean,
): Promise<OcrVisionResult> {
  if (requestedMode === "invoice" || isLikelyInvoiceDocuments(documents)) {
    return extractInvoiceFieldsFromDocuments(documents, remoteEnabled);
  }
  return buildImageOcrResult(documents);
}

export function emptyOcrResult(mode: "image" | "invoice", errors: string[] = []): OcrVisionResult {
  return {
    kind: mode,
    text: "",
    rows: [],
    fields: {},
    invoices: [],
    errors,
  };
}

export function formatOcrDocumentErrors(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.error)
    .map((document) => `${document.filename}: ${document.error}`)
    .join("\n");
}

export function normalizeOcrVisionResult(
  mode: "image" | "invoice",
  content: string,
): OcrVisionResult {
  const parsed = parseJsonObject(content);
  if (!parsed) return normalizePlainOcrText(mode, content);

  const invoices = normalizeInvoices(parsed.invoices);
  return {
    kind: mode,
    text:
      typeof parsed.text === "string"
        ? parsed.text
        : invoices
            .map((item) => item.text)
            .filter(Boolean)
            .join("\n"),
    rows: normalizeRows(parsed.rows),
    fields: normalizeFields(parsed.fields),
    invoices,
    errors: normalizeStringArray(parsed.errors),
  };
}

function isLikelyInvoiceDocuments(documents: MineruParsedDocument[]): boolean {
  return documents.some((document) =>
    isLikelyInvoiceText(`${document.filename}\n${document.text}`),
  );
}

function buildImageOcrResult(documents: MineruParsedDocument[]): OcrVisionResult {
  return {
    kind: "image",
    text: buildCombinedText(documents),
    rows: documents.flatMap((document) => document.rows),
    fields: {},
    invoices: [],
    errors: normalizeStringArray(formatOcrDocumentErrors(documents).split("\n")),
  };
}

async function extractInvoiceFieldsFromDocuments(
  documents: MineruParsedDocument[],
  remoteEnabled: boolean,
): Promise<OcrVisionResult> {
  const fallback = buildInvoiceFallbackResult(documents);
  if (!remoteEnabled) return fallback;

  try {
    assertRemoteDataProcessingAllowed({
      enabled: true,
      operation: "invoice-extraction",
      texts: documents.map((document) => document.text),
    });
    const aiConfig = getActiveAIConfig();
    const aiClient = createAIClient(aiConfig);
    const result = await aiClient.chat({
      messages: [{ role: "user", content: buildInvoicePrompt(documents) }],
      maxTokens: 4000,
      temperature: 0,
      reasoningMode: "off",
    });
    const normalized = normalizeOcrVisionResult("invoice", result.content || "");
    const extractedInvoices = normalizeExtractedInvoiceItems(normalized, documents);
    const mergedInvoices = mergeInvoiceItems(extractedInvoices, documents);
    return {
      kind: "invoice",
      text: normalized.text || fallback.text,
      rows: normalized.rows.length > 0 ? normalized.rows : fallback.rows,
      fields: Object.keys(normalized.fields).length > 0 ? normalized.fields : fallback.fields,
      invoices: mergedInvoices.length > 0 ? mergedInvoices : fallback.invoices,
      errors: [...fallback.errors, ...normalized.errors],
      remoteProcessing: [
        {
          operation: "invoice-extraction",
          service: aiConfig.provider,
          destination: getDestinationHost(aiConfig.baseUrl),
          dataSummary: `${documents.length} 个文件的 OCR 文本`,
        },
      ],
    };
  } catch (error: unknown) {
    const message =
      error && typeof error === "object" && "message" in error && typeof error.message === "string"
        ? error.message
        : "未知错误";
    return {
      ...fallback,
      errors: [...fallback.errors, `发票字段抽取失败，已保留 MinerU OCR 文本：${message}`],
    };
  }
}

function getDestinationHost(baseUrl: string): string {
  try {
    return new URL(baseUrl).host;
  } catch {
    return baseUrl;
  }
}

function normalizeExtractedInvoiceItems(
  normalized: OcrVisionResult,
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  if (normalized.invoices.length > 0) return normalized.invoices;
  if (Object.keys(normalized.fields).length === 0) return [];
  const document = documents[0];
  return [
    {
      filename: document?.filename || "识别文本",
      text: normalized.text || document?.text || "",
      fields: normalized.fields,
      rows: normalized.rows,
    },
  ];
}

function buildInvoiceFallbackResult(documents: MineruParsedDocument[]): OcrVisionResult {
  const extracted = buildInvoiceFieldFallback(documents);
  return {
    kind: "invoice",
    text: buildCombinedText(documents),
    rows:
      extracted.rows.length > 0 ? extracted.rows : documents.flatMap((document) => document.rows),
    fields: extracted.fields,
    invoices: extracted.invoices,
    errors: normalizeStringArray(formatOcrDocumentErrors(documents).split("\n")),
  };
}

function buildInvoicePrompt(documents: MineruParsedDocument[]): string {
  return [
    "下面是 MinerU 通用 OCR/版面解析得到的发票 Markdown 文本。",
    "请只基于这些文本抽取发票字段，不要编造缺失信息。",
    "只返回严格 JSON，不要 Markdown，不要解释。",
    "JSON 结构必须是：",
    '{"kind":"invoice","text":"合并后的可读文本","rows":[["列1","列2"]],"fields":{"字段":"值"},"invoices":[{"filename":"文件名","text":"文本","fields":{"发票号码":"","开票日期":"","购买方名称":"","销售方名称":"","金额":"","税额":"","价税合计":""},"rows":[["列1","列2"]]}],"errors":[]}',
    "字段优先包含：发票号码、开票日期、购买方名称、购买方税号、销售方名称、销售方税号、金额、税额、价税合计、发票类型、校验码、备注。",
    "每个输入文件都要在 invoices 中返回一项；未识别字段填空字符串。",
    "",
    buildLimitedSource(documents),
  ].join("\n");
}

function buildCombinedText(documents: MineruParsedDocument[]): string {
  return documents
    .filter((document) => document.text.trim())
    .map((document) => `## ${document.filename}\n${document.text.trim()}`)
    .join("\n\n");
}

function buildLimitedSource(documents: MineruParsedDocument[]): string {
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

function mergeInvoiceItems(
  extractedInvoices: OcrInvoiceItem[],
  documents: MineruParsedDocument[],
): OcrInvoiceItem[] {
  const fallbackInvoices = buildInvoiceFieldFallback(documents).invoices;
  if (documents.length === 0) return extractedInvoices;
  return documents.map((document, index) => {
    const extracted =
      extractedInvoices.find((invoice) => invoice.filename === document.filename) ||
      extractedInvoices[index];
    const fallback =
      fallbackInvoices.find((invoice) => invoice.filename === document.filename) ||
      fallbackInvoices[index];
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

function parseJsonObject(content: string): Record<string, unknown> | null {
  const trimmed = content.trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  const candidate = fenced ? fenced[1].trim() : trimmed;
  const start = candidate.indexOf("{");
  const end = candidate.lastIndexOf("}");
  if (start < 0 || end <= start) return null;
  try {
    const parsed: unknown = JSON.parse(candidate.slice(start, end + 1));
    return parsed && typeof parsed === "object" && !Array.isArray(parsed)
      ? (parsed as Record<string, unknown>)
      : null;
  } catch {
    return null;
  }
}

function normalizeInvoices(value: unknown): OcrInvoiceItem[] {
  if (!Array.isArray(value)) return [];
  return value.flatMap((item): OcrInvoiceItem[] => {
    if (!item || typeof item !== "object" || Array.isArray(item)) return [];
    const raw = item as Record<string, unknown>;
    return [
      {
        filename: typeof raw.filename === "string" ? raw.filename : "",
        text: typeof raw.text === "string" ? raw.text : "",
        fields: normalizeFields(raw.fields),
        rows: normalizeRows(raw.rows),
        error: typeof raw.error === "string" ? raw.error : undefined,
      },
    ];
  });
}

function normalizeFields(value: unknown): Record<string, string> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const fields: Record<string, string> = {};
  for (const [key, rawValue] of Object.entries(value)) {
    if (!key.trim()) continue;
    fields[key] =
      rawValue === null || rawValue === undefined
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
    return [row.map((cell) => (cell === null || cell === undefined ? "" : String(cell)))];
  });
}

function normalizeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => item.trim())
    .filter(Boolean);
}
