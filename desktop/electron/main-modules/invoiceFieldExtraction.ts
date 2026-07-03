import type { MineruParsedDocument } from "./mineruOcr";

export interface ExtractedInvoiceItem {
  filename: string;
  text: string;
  fields: Record<string, string>;
  rows: string[][];
  error?: string;
}

export interface ExtractedInvoiceFallback {
  fields: Record<string, string>;
  invoices: ExtractedInvoiceItem[];
  rows: string[][];
}

const INVOICE_FIELD_ORDER = [
  "文件名",
  "发票号码",
  "开票日期",
  "购买方名称",
  "购买方税号",
  "销售方名称",
  "销售方税号",
  "金额",
  "税额",
  "价税合计",
  "发票类型",
  "校验码",
  "备注",
];

const INVOICE_ALIASES: Record<string, string[]> = {
  发票号码: ["发票号码", "发票号", "票据号码"],
  开票日期: ["开票日期", "发票日期", "日期"],
  购买方名称: ["购买方名称", "购方名称", "购买方", "购方"],
  购买方税号: ["购买方税号", "购买方纳税人识别号", "购买方统一社会信用代码", "购方税号", "购方纳税人识别号", "购方统一社会信用代码"],
  销售方名称: ["销售方名称", "销方名称", "销售方", "销方"],
  销售方税号: ["销售方税号", "销售方纳税人识别号", "销售方统一社会信用代码", "销方税号", "销方纳税人识别号", "销方统一社会信用代码"],
  金额: ["不含税金额", "合计金额", "金额"],
  税额: ["合计税额", "税额"],
  价税合计: ["价税合计", "小写金额", "小写", "合计"],
  发票类型: ["发票类型", "票种", "类型"],
  校验码: ["校验码"],
  备注: ["备注"],
};

const ALL_ALIASES = Object.values(INVOICE_ALIASES)
  .flat()
  .sort((left, right) => right.length - left.length);

type PartyContext = "buyer" | "seller" | null;

export function isLikelyInvoiceText(text: string): boolean {
  const normalized = normalizeInvoiceLabel(text);
  if (!normalized) return false;
  const hits = [
    /发票/.test(normalized),
    /发票号码|发票号|票据号码/.test(normalized),
    /开票日期|发票日期/.test(normalized),
    /购买方|购方/.test(normalized),
    /销售方|销方/.test(normalized),
    /价税合计|合计金额|合计税额|校验码/.test(normalized),
  ].filter(Boolean).length;
  return hits >= 2;
}

export function buildInvoiceFieldFallback(
  documents: MineruParsedDocument[],
): ExtractedInvoiceFallback {
  const invoices = documents.map((document) => {
    const fields = completeInvoiceFields({
      文件名: document.filename,
      ...extractInvoiceFields(document),
    });
    return {
      filename: document.filename,
      text: document.text,
      fields,
      rows: Object.keys(fields).length > 0
        ? buildRowsFromFields([fields])
        : document.rows,
      error: document.error,
    };
  });

  const fields = mergeInvoiceFields(...invoices.map((invoice) => invoice.fields));
  return {
    fields,
    invoices,
    rows: Object.keys(fields).length > 0
      ? buildRowsFromFields(invoices.map((invoice) => invoice.fields), Object.keys(fields))
      : documents.flatMap((document) => document.rows),
  };
}

export function buildRowsFromFields(
  fieldRows: Record<string, string>[],
  fieldOrder = INVOICE_FIELD_ORDER,
): string[][] {
  const extraFields = fieldRows.flatMap((fields) => Object.keys(fields))
    .filter((field) => !fieldOrder.includes(field));
  const headers = [...fieldOrder, ...Array.from(new Set(extraFields))]
    .filter((field) => fieldRows.some((fields) => Object.prototype.hasOwnProperty.call(fields, field)));

  if (headers.length === 0) return [];
  return [
    headers,
    ...fieldRows.map((fields) => headers.map((field) => fields[field] ?? "")),
  ];
}

export function mergeInvoiceFields(...sources: Array<Record<string, string> | undefined>): Record<string, string> {
  const merged: Record<string, string> = {};
  for (const source of sources) {
    if (!source) continue;
    for (const [field, rawValue] of Object.entries(source)) {
      const value = cleanInvoiceValue(rawValue);
      if (!field.trim()) continue;
      if (!Object.prototype.hasOwnProperty.call(merged, field) || value) {
        merged[field] = value;
      }
    }
  }
  return merged;
}

function extractInvoiceFields(document: MineruParsedDocument): Record<string, string> {
  const fields: Record<string, string> = {};
  extractFromRows(document.rows, fields);
  extractFromText(document.text, fields);
  return fields;
}

function extractFromRows(rows: string[][], fields: Record<string, string>): void {
  let context: PartyContext = null;

  for (const row of rows) {
    const cells = row.map((cell) => cleanInvoiceValue(cell)).filter(Boolean);
    if (cells.length === 0) continue;

    context = updatePartyContext(cells.join(" "), context);
    if (cells.length >= 2) {
      for (let index = 0; index < cells.length - 1; index += 1) {
        const field = matchInvoiceField(cells[index], context);
        if (field) {
          assignField(fields, field, cells[index + 1]);
          index += 1;
        }
      }
    }

    if (cells.length === 2) {
      const field = matchInvoiceField(cells[0], context);
      if (field) assignField(fields, field, cells[1]);
    }
  }
}

function extractFromText(text: string, fields: Record<string, string>): void {
  const lines = text.split(/\r?\n/).map((line) => cleanInvoiceValue(line)).filter(Boolean);
  let context: PartyContext = null;

  for (const line of lines) {
    context = updatePartyContext(line, context);
    for (const [field, aliases] of Object.entries(INVOICE_ALIASES)) {
      if (fields[field]) continue;
      for (const alias of aliases) {
        const value = extractValueAfterAlias(line, alias);
        if (!value) continue;
        const contextualField = matchInvoiceField(alias, context) || field;
        assignField(fields, contextualField, value);
        break;
      }
    }
  }
}

function extractValueAfterAlias(line: string, alias: string): string {
  const index = line.indexOf(alias);
  if (index < 0) return "";

  let value = line.slice(index + alias.length);
  value = value.replace(/^（[^）]*）/, "");
  value = value.replace(/^\([^)]*\)/, "");
  value = value.replace(/^[\s:：|｜\-—_]+/, "");
  value = trimAtNextAlias(value);
  return cleanInvoiceValue(value);
}

function trimAtNextAlias(value: string): string {
  let end = value.length;
  for (const alias of ALL_ALIASES) {
    const index = value.indexOf(alias);
    if (index > 0 && index < end) end = index;
  }
  return value.slice(0, end);
}

function assignField(fields: Record<string, string>, field: string, rawValue: string): void {
  const value = cleanInvoiceValue(rawValue);
  if (!value || looksLikeFieldLabelOnly(value)) return;
  if (!fields[field]) fields[field] = value;
}

function completeInvoiceFields(fields: Record<string, string>): Record<string, string> {
  if (!Object.values(fields).some(Boolean)) return {};
  const completed: Record<string, string> = {};
  for (const field of INVOICE_FIELD_ORDER) {
    completed[field] = fields[field] ?? "";
  }
  for (const [field, value] of Object.entries(fields)) {
    completed[field] = value;
  }
  return completed;
}

function matchInvoiceField(label: string, context: PartyContext): string {
  const normalized = normalizeInvoiceLabel(label);
  if (!normalized) return "";

  if (/购买方|购方/.test(label)) {
    if (/税号|纳税人识别号|统一社会信用代码/.test(label)) return "购买方税号";
    if (/名称|名\s*称/.test(label) || /购买方|购方/.test(label)) return "购买方名称";
  }
  if (/销售方|销方/.test(label)) {
    if (/税号|纳税人识别号|统一社会信用代码/.test(label)) return "销售方税号";
    if (/名称|名\s*称/.test(label) || /销售方|销方/.test(label)) return "销售方名称";
  }
  if (/税号|纳税人识别号|统一社会信用代码/.test(label)) {
    if (context === "buyer") return "购买方税号";
    if (context === "seller") return "销售方税号";
  }
  if (/名称|名\s*称/.test(label)) {
    if (context === "buyer") return "购买方名称";
    if (context === "seller") return "销售方名称";
  }

  for (const [field, aliases] of Object.entries(INVOICE_ALIASES)) {
    if (aliases.some((alias) => normalized.includes(normalizeInvoiceLabel(alias)))) {
      return field;
    }
  }
  return "";
}

function updatePartyContext(text: string, current: PartyContext): PartyContext {
  if (/购买方|购方/.test(text)) return "buyer";
  if (/销售方|销方/.test(text)) return "seller";
  return current;
}

function cleanInvoiceValue(value: unknown): string {
  return String(value ?? "")
    .replace(/<[^>]+>/g, " ")
    .replace(/[*`#]/g, "")
    .replace(/[|｜]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/^[\s:：,，;；]+|[\s,，;；]+$/g, "")
    .trim();
}

function normalizeInvoiceLabel(value: string): string {
  return cleanInvoiceValue(value).replace(/[()\[\]（）【】\s:：,，;；/\\]/g, "");
}

function looksLikeFieldLabelOnly(value: string): boolean {
  const normalized = normalizeInvoiceLabel(value);
  return Object.values(INVOICE_ALIASES).some((aliases) =>
    aliases.some((alias) => normalized === normalizeInvoiceLabel(alias))
  );
}
