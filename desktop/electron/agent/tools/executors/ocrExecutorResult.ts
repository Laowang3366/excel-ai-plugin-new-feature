import * as path from "node:path";

import type { RemoteDataTransferSummary } from "../../../shared/egressPolicy";
import type { MineruParsedDocument } from "../../../main-modules/mineruOcr";
import type { LocalParsedDocument } from "./localDocumentParser";

export type OcrMode = "ocr" | "invoice" | "layout" | "style";
export type OcrProvider = "mineru" | "mineru-agent" | "local";

export interface OcrFallbackAttempt {
  provider: OcrProvider;
  success: boolean;
  skipped?: boolean;
  parsedFiles?: number;
  totalFiles?: number;
  reason?: string;
  error?: string;
  quotaLikely?: boolean;
}

export interface SelectedOcrDocument {
  document: MineruParsedDocument | LocalParsedDocument;
  provider: OcrProvider;
}

interface OutputDocument {
  filename: string;
  text: string;
  textTruncated: boolean;
  rows: string[][];
  rowsTruncated: boolean;
  provider: OcrProvider;
  sourceType?: string;
  warnings?: string[];
  error?: string;
}

interface BuildOcrToolResultInput {
  fallbacks: OcrFallbackAttempt[];
  filePaths: string[];
  maxTableRows: number;
  maxTextChars: number;
  mode: OcrMode;
  selected: Array<SelectedOcrDocument | undefined>;
  warnings: string[];
}

export function buildOcrToolResult({
  fallbacks,
  filePaths,
  maxTableRows,
  maxTextChars,
  mode,
  selected,
  warnings,
}: BuildOcrToolResultInput) {
  const outputDocuments = selected.flatMap((entry, index) =>
    entry
      ? [
          formatOutputDocument(
            entry.document,
            entry.provider,
            filePaths[index],
            maxTextChars,
            maxTableRows,
          ),
        ]
      : [],
  );
  const selectedProviders = new Set(outputDocuments.map((document) => document.provider));
  const selectedProvider =
    selectedProviders.size === 1 ? Array.from(selectedProviders)[0] : "mixed";
  const combinedText = buildCombinedText(outputDocuments, maxTextChars);
  const allRows = outputDocuments.flatMap((document) => document.rows);
  const rows = allRows.slice(0, maxTableRows);
  const errors = outputDocuments
    .filter((document) => document.error)
    .map((document) => `${document.filename}: ${document.error}`);
  const completedCount = outputDocuments.filter(
    (document) => document.text.trim().length > 0 || document.rows.length > 0,
  ).length;
  const complete =
    outputDocuments.length === filePaths.length && completedCount === filePaths.length;

  return {
    success: complete,
    ...(complete ? {} : { error: `OCR 仅完成 ${completedCount}/${filePaths.length} 个文件` }),
    data: {
      provider: selectedProvider,
      status: complete ? "complete" : "partial",
      mode,
      fileCount: filePaths.length,
      text: combinedText.text,
      textTruncated:
        combinedText.truncated || outputDocuments.some((document) => document.textTruncated),
      rows,
      rowsTruncated:
        allRows.length > rows.length || outputDocuments.some((document) => document.rowsTruncated),
      documents: outputDocuments,
      errors,
      warnings: normalizeWarnings(warnings, selectedProvider, mode, filePaths),
      fallbacks,
      remoteProcessing: buildOcrRemoteProcessing(outputDocuments),
      nextTools: buildSelfFallbackTools(filePaths, mode),
    },
  };
}

function formatOutputDocument(
  document: MineruParsedDocument | LocalParsedDocument,
  provider: OcrProvider,
  filePath: string,
  maxTextChars: number,
  maxTableRows: number,
): OutputDocument {
  const perDocumentMaxChars = Math.min(12_000, maxTextChars);
  const clipped = clipText(document.text, perDocumentMaxChars);
  const rowLimit = Math.min(50, maxTableRows);
  return {
    filename: document.filename,
    text: clipped.text,
    textTruncated: clipped.truncated,
    rows: document.rows.slice(0, rowLimit),
    rowsTruncated: document.rows.length > rowLimit,
    provider,
    sourceType:
      "sourceType" in document && typeof document.sourceType === "string"
        ? document.sourceType
        : path.extname(filePath).replace(/^\./, "").toLowerCase(),
    warnings: "warnings" in document ? document.warnings : undefined,
    error: document.error,
  };
}

function buildCombinedText(
  documents: Array<{ filename: string; text: string }>,
  maxChars: number,
): { text: string; truncated: boolean } {
  return clipText(
    documents
      .filter((document) => document.text.trim())
      .map((document) => `## ${document.filename}\n${document.text.trim()}`)
      .join("\n\n"),
    maxChars,
  );
}

function normalizeWarnings(
  warnings: string[],
  provider: OcrProvider | "mixed",
  mode: OcrMode,
  filePaths: string[],
): string[] {
  const normalized = Array.from(new Set(warnings.filter(Boolean)));
  if (provider === "mineru-agent") {
    normalized.push(
      "当前结果来自 MinerU 免费 Agent 轻量解析，受单文件大小、页数和 IP 限频限制；如信息不足，可稍后重试标准 MinerU 或继续用本地工具兜底。",
    );
  }
  if (provider === "local") {
    normalized.push(
      "当前结果来自本地免费兜底解析，图片/PDF 扫描件不会获得真正 OCR 文本；可结合 Office/文件工具继续检查结构或请求用户提供可复制文本。",
    );
  }
  if (mode === "layout" || mode === "style") {
    const hasOfficeFile = filePaths.some((filePath) =>
      [".docx", ".pptx", ".xlsx", ".xlsm"].includes(path.extname(filePath).toLowerCase()),
    );
    if (hasOfficeFile) {
      normalized.push(
        "文件可见内容解析主要提供文本、表格和结构线索；版面或样式验收可继续调用 office.action.inspect / office.action.validate / snapshot 等内置工具。",
      );
    }
  }
  return normalized;
}

function buildOcrRemoteProcessing(documents: OutputDocument[]): RemoteDataTransferSummary[] {
  const providers = new Set(documents.map((document) => document.provider));
  const transfers: RemoteDataTransferSummary[] = [];
  if (providers.has("mineru")) {
    transfers.push({
      operation: "ocr",
      service: "MinerU",
      destination: "mineru.net",
      dataSummary: `${documents.filter((document) => document.provider === "mineru").length} 个文件`,
    });
  }
  if (providers.has("mineru-agent")) {
    transfers.push({
      operation: "ocr",
      service: "MinerU Agent",
      destination: "mineru.net",
      dataSummary: `${documents.filter((document) => document.provider === "mineru-agent").length} 个文件`,
    });
  }
  return transfers;
}

function buildSelfFallbackTools(filePaths: string[], mode: OcrMode): Array<Record<string, string>> {
  const extensions = new Set(filePaths.map((filePath) => path.extname(filePath).toLowerCase()));
  const tools: Array<Record<string, string>> = [];

  if ([...extensions].some((ext) => [".xlsx", ".xlsm", ".docx", ".pptx"].includes(ext))) {
    tools.push({
      tool: "office.action.inspect",
      useWhen: "需要继续读取 Office 文件结构、文本部件、表格、布局对象或修改后的文件状态",
    });
    tools.push({
      tool: "office.action.validate",
      useWhen: "完成文件级修改后验证数量、结构、样式或对象变化",
    });
  }
  if (mode === "layout" || mode === "style") {
    tools.push({
      tool: "office.action.inspect",
      useWhen: "需要从 Office 文件自身结构继续判断版面、表格和样式状态",
    });
  }
  tools.push({
    tool: "file.getPaths",
    useWhen: "用户只给出桌面、下载、文档等模糊位置时，先解析成本机绝对路径",
  });
  return dedupeToolSuggestions(tools);
}

function dedupeToolSuggestions(
  tools: Array<Record<string, string>>,
): Array<Record<string, string>> {
  const seen = new Set<string>();
  return tools.filter((item) => {
    const key = `${item.tool}:${item.useWhen}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function clipText(value: string, maxChars: number): { text: string; truncated: boolean } {
  if (value.length <= maxChars) return { text: value, truncated: false };
  return { text: value.slice(0, maxChars), truncated: true };
}
