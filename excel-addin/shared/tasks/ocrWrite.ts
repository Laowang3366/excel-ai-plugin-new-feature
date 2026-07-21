/** OCR result helpers for preview / sheet write (desktop-aligned, no IPC). */

export interface OcrInvoiceItem {
  filename: string;
  text: string;
  fields: Record<string, string>;
  rows: string[][];
  error?: string;
}

export interface OcrResult {
  kind: "image" | "invoice";
  text: string;
  rows: string[][];
  fields: Record<string, string>;
  invoices: OcrInvoiceItem[];
  errors: string[];
}

export function extractOcrFieldNames(result: OcrResult): string[] {
  if (result.kind === "invoice" && result.invoices.length > 0) {
    return Array.from(
      new Set(result.invoices.flatMap((invoice) => Object.keys(invoice.fields || {}))),
    );
  }
  if (result.rows.length > 0) return result.rows[0] ?? [];
  if (result.fields && Object.keys(result.fields).length > 0) {
    return Object.keys(result.fields);
  }
  if (result.invoices.length > 0) {
    return Array.from(
      new Set(result.invoices.flatMap((invoice) => Object.keys(invoice.fields || {}))),
    );
  }
  return [];
}

export function canWriteOcrResult(result: OcrResult, selectedFields: string[]): boolean {
  return (
    selectedFields.length > 0 ||
    (extractOcrFieldNames(result).length === 0 && Boolean(result.text.trim()))
  );
}

export function buildOcrWriteValues(result: OcrResult, selectedFields: string[]): string[][] {
  if (selectedFields.length === 0) {
    return result.text ? [[result.text]] : [];
  }
  if (result.kind === "invoice" && result.invoices.length > 0) {
    const rows = result.invoices.map((invoice) =>
      selectedFields.map((field) => stringifyCell(invoice.fields[field] ?? "")),
    );
    return [selectedFields, ...rows];
  }
  if (result.rows.length > 0) {
    return filterRowsForFields(result.rows, selectedFields);
  }
  if (Object.keys(result.fields).length > 0) {
    return [
      selectedFields,
      selectedFields.map((field) => stringifyCell(result.fields[field] ?? "")),
    ];
  }
  return [[result.text || ""]];
}

function stringifyCell(value: unknown): string {
  return value === null || value === undefined ? "" : String(value);
}

function filterRowsForFields(rows: string[][], selectedFields: string[]): string[][] {
  if (rows.length === 0 || selectedFields.length === 0) return rows;
  const header = rows[0] ?? [];
  const colIndices = selectedFields
    .map((field) => header.indexOf(field))
    .filter((index) => index >= 0);
  if (colIndices.length === 0) return rows;
  return rows.map((row) => colIndices.map((index) => row[index] ?? ""));
}
