import * as path from "path";
import type { MineruParsedDocument } from "./mineruOcr";
import { buildInvoiceFieldFallback, isLikelyInvoiceText } from "./invoiceFieldExtraction";
import type { RemoteDataTransferSummary } from "../shared/egressPolicy";

export interface OcrInvoiceItem {
  filename: string;
  text: string;
  fields: Record<string, string>;
  rows: string[][];
  error?: string;
}

export interface OcrVisionResult {
  kind: "image" | "invoice";
  text: string;
  rows: string[][];
  fields: Record<string, string>;
  invoices: OcrInvoiceItem[];
  errors: string[];
  remoteProcessing?: RemoteDataTransferSummary[];
}

export function normalizeOcrMode(mode: unknown): "image" | "invoice" {
  return mode === "invoice" ? "invoice" : "image";
}

export function isLikelyInvoiceFileList(filePaths: string[]): boolean {
  return filePaths.some((filePath) => /发票|invoice|fapiao|票据/i.test(path.basename(filePath)));
}

export function buildInvoiceResultFromPlainText(content: string, filename: string): OcrVisionResult {
  const document: MineruParsedDocument = {
    filename,
    text: content.trim(),
    rows: [],
  };
  const extracted = buildInvoiceFieldFallback([document]);
  return {
    kind: "invoice",
    text: content.trim(),
    rows: extracted.rows.length > 0 ? extracted.rows : [["识别文本"], [content.trim()]],
    fields: extracted.fields,
    invoices: extracted.invoices,
    errors: [],
  };
}

export function normalizePlainOcrText(mode: "image" | "invoice", content: string): OcrVisionResult {
  if (mode === "invoice" || isLikelyInvoiceText(content)) {
    return buildInvoiceResultFromPlainText(content, "识别文本");
  }
  return {
    kind: mode,
    text: content.trim(),
    rows: [["识别文本"], [content.trim()]],
    fields: {},
    invoices: [],
    errors: [],
  };
}
