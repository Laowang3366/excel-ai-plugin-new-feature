/** Strict OCR result marker protocol (anti prompt-injection). */

import type { OcrInvoiceItem, OcrResult } from "./ocrWrite";

/** Fence open/close — only content between these markers is treated as structured JSON. */
export const OCR_RESULT_MARKER_OPEN = "<<<WENGGE_OCR_RESULT_V1";
export const OCR_RESULT_MARKER_CLOSE = "WENGGE_OCR_RESULT_V1>>>";

const MAX_JSON_CHARS = 120_000;
const MAX_TEXT_CHARS = 40_000;
const MAX_ROWS = 200;
const MAX_COLS = 40;
const MAX_INVOICES = 20;
const MAX_FIELD_KEYS = 40;

export type OcrParseOk = {
  ok: true;
  result: OcrResult;
  /** Assistant prose outside the marker (may be empty). */
  narrative: string;
  rawMarkerJson: string;
};

export type OcrParseFail = {
  ok: false;
  reason: string;
  /** Full assistant text for fallback preview (never treat as structured). */
  rawText: string;
  narrative?: string;
};

export type OcrParseOutcome = OcrParseOk | OcrParseFail;

/**
 * Extract and validate a single OCR result fence from assistant text.
 * Fail closed: malformed / extra fences / unknown shape → not structured.
 */
export function parseOcrAssistantResult(assistantText: string): OcrParseOutcome {
  const rawText = typeof assistantText === "string" ? assistantText : "";
  if (!rawText.trim()) {
    return { ok: false, reason: "助手回复为空", rawText };
  }

  const open = OCR_RESULT_MARKER_OPEN;
  const close = OCR_RESULT_MARKER_CLOSE;
  const firstOpen = rawText.indexOf(open);
  if (firstOpen < 0) {
    return {
      ok: false,
      reason: "未找到结构化 OCR 标记（仍可预览原始文本）",
      rawText,
    };
  }
  const secondOpen = rawText.indexOf(open, firstOpen + open.length);
  if (secondOpen >= 0) {
    return {
      ok: false,
      reason: "检测到多个 OCR 标记，拒绝结构化解析",
      rawText,
    };
  }
  const closeIdx = rawText.indexOf(close, firstOpen + open.length);
  if (closeIdx < 0) {
    return {
      ok: false,
      reason: "OCR 标记未闭合",
      rawText,
    };
  }
  if (rawText.indexOf(close, closeIdx + close.length) >= 0) {
    return {
      ok: false,
      reason: "检测到多余的 OCR 闭合标记",
      rawText,
    };
  }

  const jsonRaw = rawText.slice(firstOpen + open.length, closeIdx).trim();
  // Strip optional surrounding code fences accidentally wrapped inside
  const jsonBody = jsonRaw
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  if (!jsonBody) {
    return { ok: false, reason: "OCR 标记内无 JSON", rawText };
  }
  if (jsonBody.length > MAX_JSON_CHARS) {
    return { ok: false, reason: "OCR JSON 过大", rawText };
  }
  if (jsonBody.includes(open) || jsonBody.includes(close)) {
    return { ok: false, reason: "OCR JSON 内含嵌套标记", rawText };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonBody);
  } catch {
    return { ok: false, reason: "OCR JSON 解析失败", rawText };
  }

  const normalized = normalizeOcrResult(parsed);
  if (!normalized.ok) {
    return { ok: false, reason: normalized.reason, rawText };
  }

  const narrative = (
    rawText.slice(0, firstOpen) + rawText.slice(closeIdx + close.length)
  )
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  return {
    ok: true,
    result: normalized.result,
    narrative,
    rawMarkerJson: jsonBody,
  };
}

function normalizeOcrResult(
  value: unknown,
): { ok: true; result: OcrResult } | { ok: false; reason: string } {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return { ok: false, reason: "OCR 根节点必须是对象" };
  }
  const obj = value as Record<string, unknown>;
  const kindRaw = obj.kind;
  if (kindRaw !== "image" && kindRaw !== "invoice") {
    return { ok: false, reason: "kind 必须是 image 或 invoice" };
  }
  const kind = kindRaw as "image" | "invoice";

  const text = clampString(obj.text, MAX_TEXT_CHARS);
  const errors = (asStringArray(obj.errors) ?? []).slice(0, 20);
  const fields = asStringRecord(obj.fields);
  const rows = asStringGrid(obj.rows);
  const invoices = asInvoices(obj.invoices);

  if (fields && Object.keys(fields).length > MAX_FIELD_KEYS) {
    return { ok: false, reason: "fields 键过多" };
  }
  if (rows && (rows.length > MAX_ROWS || rows.some((r) => r.length > MAX_COLS))) {
    return { ok: false, reason: "rows 尺寸超限" };
  }
  if (invoices && invoices.length > MAX_INVOICES) {
    return { ok: false, reason: "invoices 数量超限" };
  }

  const result: OcrResult = {
    kind,
    text: text ?? "",
    rows: rows ?? [],
    fields: fields ?? {},
    invoices: invoices ?? [],
    errors: errors ?? [],
  };

  // At least some content
  if (
    !result.text.trim() &&
    result.rows.length === 0 &&
    Object.keys(result.fields).length === 0 &&
    result.invoices.length === 0
  ) {
    return { ok: false, reason: "结构化结果无有效内容" };
  }

  return { ok: true, result };
}

function clampString(v: unknown, max: number): string | undefined {
  if (v === undefined || v === null) return undefined;
  if (typeof v !== "string") return undefined;
  return v.length > max ? v.slice(0, max) : v;
}

function asStringArray(v: unknown): string[] | undefined {
  if (v === undefined) return undefined;
  if (!Array.isArray(v)) return undefined;
  const out: string[] = [];
  for (const item of v) {
    if (typeof item === "string") out.push(item.slice(0, 500));
    else if (item == null) out.push("");
    else out.push(String(item).slice(0, 500));
  }
  return out;
}

function asStringRecord(v: unknown): Record<string, string> | undefined {
  if (v === undefined) return undefined;
  if (!v || typeof v !== "object" || Array.isArray(v)) return undefined;
  const out: Record<string, string> = {};
  for (const [k, val] of Object.entries(v as Record<string, unknown>)) {
    if (typeof k !== "string" || !k.trim()) continue;
    if (k.length > 80) continue;
    out[k.slice(0, 80)] =
      val == null ? "" : String(val).slice(0, MAX_TEXT_CHARS);
  }
  return out;
}

function asStringGrid(v: unknown): string[][] | undefined {
  if (v === undefined) return undefined;
  if (!Array.isArray(v)) return undefined;
  const rows: string[][] = [];
  for (const row of v) {
    if (!Array.isArray(row)) return undefined;
    rows.push(
      row.map((cell) =>
        cell == null ? "" : String(cell).slice(0, 2000),
      ),
    );
  }
  return rows;
}

function asInvoices(v: unknown): OcrInvoiceItem[] | undefined {
  if (v === undefined) return undefined;
  if (!Array.isArray(v)) return undefined;
  const out: OcrInvoiceItem[] = [];
  for (const item of v) {
    if (!item || typeof item !== "object" || Array.isArray(item)) continue;
    const rec = item as Record<string, unknown>;
    out.push({
      filename: clampString(rec.filename, 260) ?? "",
      text: clampString(rec.text, MAX_TEXT_CHARS) ?? "",
      fields: asStringRecord(rec.fields) ?? {},
      rows: asStringGrid(rec.rows) ?? [],
      error: clampString(rec.error, 500),
    });
  }
  return out;
}

/** Build instruction block embedded in OCR task payload / scenario. */
export function ocrResultProtocolInstruction(): string {
  return [
    "结构化输出协议（必须遵守）：",
    `在回复末尾输出且仅输出一处标记块：以 ${OCR_RESULT_MARKER_OPEN} 单独起行，下一行起为单个 JSON 对象，最后以 ${OCR_RESULT_MARKER_CLOSE} 单独结束。`,
    'JSON 字段：kind("image"|"invoice")、text(string)、fields(object string→string)、rows(string[][])、invoices(array)、errors(string[])。',
    "标记外可写简短说明；禁止在标记内嵌套相同标记；禁止把 Base64/API Key 写入 text 或 fields。",
    "加载项 UI 仅信任该标记内 JSON；解析失败时用户将只看到原始文本。",
  ].join("\n");
}

/**
 * Parse write target `Sheet1!A1` / `'My Sheet'!B2:C3`.
 * Bare A1 without sheet returns null (caller must require full address).
 */
export function parseSheetRangeAddress(
  address: string,
): { sheetName: string; range: string } | null {
  const raw = (address ?? "").trim();
  if (!raw) return null;
  // Quoted sheet: 'Sheet Name'!A1
  const quoted = /^'((?:[^']|'')+)'!(.+)$/s.exec(raw);
  if (quoted) {
    const sheetName = (quoted[1] ?? "").replace(/''/g, "'").trim();
    const range = (quoted[2] ?? "").trim();
    if (!sheetName || !range || /!/.test(range)) return null;
    return { sheetName, range };
  }
  // Unquoted: Sheet1!A1 (sheet cannot contain !)
  const plain = /^([^'!]+)!(.+)$/s.exec(raw);
  if (plain) {
    const sheetName = (plain[1] ?? "").trim();
    const range = (plain[2] ?? "").trim();
    if (!sheetName || !range || /!/.test(range)) return null;
    return { sheetName, range };
  }
  return null;
}

/** Redact accidental secrets from text shown in OCR UI. */
export function sanitizeOcrUiText(text: string): string {
  if (!text) return "";
  return text
    .replace(/sk-[a-zA-Z0-9]{10,}/g, "[redacted-key]")
    .replace(
      /data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=\s]{40,}/g,
      "[omitted image data]",
    )
    .replace(/\b[A-Za-z0-9+/]{80,}={0,2}\b/g, "[omitted binary]");
}
