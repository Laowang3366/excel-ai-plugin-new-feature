/**
 * OCR/document visibility tool executor.
 *
 * Fallback order:
 * 1. MinerU standard API with the configured token.
 * 2. MinerU Agent lightweight API without token.
 * 3. Local free parsing and built-in Office/file tool suggestions.
 */

import * as fs from "fs";
import * as path from "path";
import type { ToolExecutor } from "../../shared/types";
import {
  parseFilesWithMineru,
  parseFilesWithMineruAgent,
  type MineruParsedDocument,
} from "../../../main-modules/mineruOcr";
import { parseFilesLocally, type LocalParsedDocument } from "./localDocumentParser";
import { validateArgs } from "./validation";

export interface OcrExecutorDeps {
  getMineruApiToken?: () => string;
}

type OcrMode = "ocr" | "invoice" | "layout" | "style";
type OcrProvider = "mineru" | "mineru-agent" | "local";

interface FallbackAttempt {
  provider: OcrProvider;
  success: boolean;
  skipped?: boolean;
  parsedFiles?: number;
  totalFiles?: number;
  reason?: string;
  error?: string;
  quotaLikely?: boolean;
}

interface OutputDocument {
  filename: string;
  text: string;
  textTruncated: boolean;
  rows: string[][];
  provider: OcrProvider;
  sourceType?: string;
  warnings?: string[];
  error?: string;
}

const SUPPORTED_EXTENSIONS = new Set([
  ".png",
  ".jpg",
  ".jpeg",
  ".webp",
  ".bmp",
  ".tif",
  ".tiff",
  ".pdf",
  ".doc",
  ".docx",
  ".ppt",
  ".pptx",
  ".xls",
  ".xlsx",
  ".xlsm",
  ".csv",
  ".md",
  ".txt",
]);

export function addOcrExecutors(target: Map<string, ToolExecutor>, deps: OcrExecutorDeps = {}): void {
  target.set("ocr.parseDocument", {
    name: "ocr.parseDocument",
    execute: async (args: Record<string, unknown>) => {
      const err = validateArgs(args, { filePaths: "array" });
      if (err) return { success: false, error: err };

      const filePaths = normalizeFilePaths(args.filePaths);
      if (filePaths.length === 0) {
        return { success: false, error: "参数 filePaths 至少需要一个有效文件路径" };
      }

      const pathError = await validateOcrFilePaths(filePaths);
      if (pathError) return { success: false, error: pathError };

      const mode = normalizeOcrToolMode(args.mode);
      const maxTextChars = clampNumber(args.maxTextChars, 60_000, 1_000, 120_000);
      const maxTableRows = clampNumber(args.maxTableRows, 200, 0, 1_000);
      const allowTokenMineru = args.allowTokenMineru !== false;
      const allowFreeMineru = args.allowFreeMineru !== false;
      const allowLocalFallback = args.allowLocalFallback !== false;
      const token = getConfiguredMineruToken(deps);

      const fallbacks: FallbackAttempt[] = [];
      const warnings: string[] = [];
      let selectedProvider: OcrProvider | undefined;
      let selectedDocuments: Array<MineruParsedDocument | LocalParsedDocument> = [];

      if (!allowTokenMineru) {
        fallbacks.push({
          provider: "mineru",
          success: false,
          skipped: true,
          reason: "调用参数 allowTokenMineru=false，已跳过配置 token 的 MinerU 标准解析",
        });
      } else if (!token) {
        fallbacks.push({
          provider: "mineru",
          success: false,
          skipped: true,
          reason: "MinerU API Token 未配置，直接尝试 MinerU 免费 Agent 轻量解析",
        });
      } else {
        const standardAttempt = await tryParseWithProvider("mineru", () => parseFilesWithMineru(filePaths, token));
        fallbacks.push(standardAttempt.fallback);
        if (standardAttempt.documents && hasAnyUsefulDocument(standardAttempt.documents)) {
          selectedProvider = "mineru";
          selectedDocuments = standardAttempt.documents;
        } else if (standardAttempt.fallback.error) {
          warnings.push(`MinerU 标准解析不可用，已进入免费 Agent 降级: ${standardAttempt.fallback.error}`);
        }
      }

      if (!selectedProvider) {
        if (!allowFreeMineru) {
          fallbacks.push({
            provider: "mineru-agent",
            success: false,
            skipped: true,
            reason: "调用参数 allowFreeMineru=false，已跳过 MinerU 免费 Agent 轻量解析",
          });
        } else {
          const agentAttempt = await tryParseWithProvider("mineru-agent", () => parseFilesWithMineruAgent(filePaths));
          fallbacks.push(agentAttempt.fallback);
          if (agentAttempt.documents && hasAnyUsefulDocument(agentAttempt.documents)) {
            selectedProvider = "mineru-agent";
            selectedDocuments = agentAttempt.documents;
          } else if (agentAttempt.fallback.error) {
            warnings.push(`MinerU 免费 Agent 解析不可用，已进入本地兜底: ${agentAttempt.fallback.error}`);
          }
        }
      }

      if (!selectedProvider) {
        if (!allowLocalFallback) {
          return {
            success: false,
            error: "MinerU 标准解析和免费 Agent 解析均不可用，且 allowLocalFallback=false，无法继续本地兜底",
            data: { fallbacks },
          };
        }
        const localDocuments = await parseFilesLocally(filePaths);
        fallbacks.push({
          provider: "local",
          success: hasAnyUsefulDocument(localDocuments),
          parsedFiles: localDocuments.filter(hasUsefulDocument).length,
          totalFiles: filePaths.length,
          reason: "远程解析不可用或未返回可用内容，已使用本地免费解析兜底",
        });
        selectedProvider = "local";
        selectedDocuments = localDocuments;
        warnings.push(...localDocuments.flatMap((document) => document.warnings));
      }

      const outputDocuments = selectedDocuments.map((document, index) =>
        formatOutputDocument(document, selectedProvider!, filePaths[index], maxTextChars, maxTableRows)
      );
      const combinedText = buildCombinedText(outputDocuments, maxTextChars);
      const allRows = outputDocuments.flatMap((document) => document.rows);
      const rows = allRows.slice(0, maxTableRows);
      const errors = outputDocuments
        .filter((document) => document.error)
        .map((document) => `${document.filename}: ${document.error}`);

      return {
        success: true,
        data: {
          provider: selectedProvider,
          mode,
          fileCount: filePaths.length,
          text: combinedText.text,
          textTruncated: combinedText.truncated,
          rows,
          rowsTruncated: allRows.length > rows.length,
          documents: outputDocuments,
          errors,
          warnings: normalizeWarnings(warnings, selectedProvider, mode, filePaths),
          fallbacks,
          nextTools: buildSelfFallbackTools(filePaths, mode, selectedProvider),
        },
      };
    },
  });
}

async function tryParseWithProvider(
  provider: OcrProvider,
  parse: () => Promise<MineruParsedDocument[]>
): Promise<{ fallback: FallbackAttempt; documents?: MineruParsedDocument[] }> {
  try {
    const documents = await parse();
    return {
      documents,
      fallback: {
        provider,
        success: hasAnyUsefulDocument(documents),
        parsedFiles: documents.filter(hasUsefulDocument).length,
        totalFiles: documents.length,
        reason: hasAnyUsefulDocument(documents) ? undefined : `${provider} 未返回可用文本或表格`,
      },
    };
  } catch (error: any) {
    const message = error?.message || String(error);
    return {
      fallback: {
        provider,
        success: false,
        error: message,
        quotaLikely: isQuotaLikeError(message),
      },
    };
  }
}

function normalizeFilePaths(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .filter((item): item is string => typeof item === "string")
    .map((item) => item.trim())
    .filter(Boolean);
}

async function validateOcrFilePaths(filePaths: string[]): Promise<string | null> {
  for (const filePath of filePaths) {
    if (!path.isAbsolute(filePath)) {
      return `OCR 文件路径必须是绝对路径: ${filePath}`;
    }
    const ext = path.extname(filePath).toLowerCase();
    if (!SUPPORTED_EXTENSIONS.has(ext)) {
      return `不支持的文件可见内容解析类型: ${ext || "未知"}，仅支持图片、PDF、Office 文档、CSV、Markdown 和纯文本`;
    }
    try {
      const stat = await fs.promises.stat(filePath);
      if (!stat.isFile()) return `OCR 路径不是文件: ${filePath}`;
    } catch {
      return `OCR 文件不存在或不可访问: ${filePath}`;
    }
  }
  return null;
}

function normalizeOcrToolMode(value: unknown): OcrMode {
  return value === "invoice" || value === "layout" || value === "style" ? value : "ocr";
}

function getConfiguredMineruToken(deps: OcrExecutorDeps): string {
  const tokenFromSettings = deps.getMineruApiToken?.().trim() || "";
  return tokenFromSettings || (process.env.MINERU_API_TOKEN || "").trim();
}

function hasUsefulDocument(document: { text: string; rows: string[][] }): boolean {
  return document.text.trim().length > 0 || document.rows.length > 0;
}

function hasAnyUsefulDocument(documents: Array<{ text: string; rows: string[][] }>): boolean {
  return documents.some(hasUsefulDocument);
}

function formatOutputDocument(
  document: MineruParsedDocument | LocalParsedDocument,
  provider: OcrProvider,
  filePath: string,
  maxTextChars: number,
  maxTableRows: number
): OutputDocument {
  const perDocumentMaxChars = Math.min(12_000, maxTextChars);
  const clipped = clipText(document.text, perDocumentMaxChars);
  return {
    filename: document.filename,
    text: clipped.text,
    textTruncated: clipped.truncated,
    rows: document.rows.slice(0, Math.min(50, maxTableRows)),
    provider,
    sourceType: "sourceType" in document && typeof document.sourceType === "string"
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
  provider: OcrProvider,
  mode: OcrMode,
  filePaths: string[]
): string[] {
  const normalized = Array.from(new Set(warnings.filter(Boolean)));
  if (provider === "mineru-agent") {
    normalized.push("当前结果来自 MinerU 免费 Agent 轻量解析，受单文件大小、页数和 IP 限频限制；如信息不足，可稍后重试标准 MinerU 或继续用本地工具兜底。");
  }
  if (provider === "local") {
    normalized.push("当前结果来自本地免费兜底解析，图片/PDF 扫描件不会获得真正 OCR 文本；可结合 Office/文件工具继续检查结构或请求用户提供可复制文本。");
  }
  if (mode === "layout" || mode === "style") {
    const hasOfficeFile = filePaths.some((filePath) =>
      [".docx", ".pptx", ".xlsx", ".xlsm"].includes(path.extname(filePath).toLowerCase())
    );
    if (hasOfficeFile) {
      normalized.push("文件可见内容解析主要提供文本、表格和结构线索；版面或样式验收可继续调用 office.action.inspect / office.action.validate / snapshot 等内置工具。");
    }
  }
  return normalized;
}

function buildSelfFallbackTools(
  filePaths: string[],
  mode: OcrMode,
  provider: OcrProvider
): Array<Record<string, string>> {
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
  if (provider === "local" && [...extensions].some((ext) => [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".pdf"].includes(ext))) {
    tools.push({
      tool: "python.execute",
      useWhen: "本地环境已有可用 OCR/转换库时，可用脚本提取文本或把文件转换成模型可读的中间格式",
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

function dedupeToolSuggestions(tools: Array<Record<string, string>>): Array<Record<string, string>> {
  const seen = new Set<string>();
  return tools.filter((item) => {
    const key = `${item.tool}:${item.useWhen}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function isQuotaLikeError(error: string): boolean {
  return /quota|limit|limited|rate|credit|balance|insufficient|次数|额度|余额|限额|频率|too many|429/i.test(error);
}

function clipText(value: string, maxChars: number): { text: string; truncated: boolean } {
  if (value.length <= maxChars) return { text: value, truncated: false };
  return { text: value.slice(0, maxChars), truncated: true };
}

function clampNumber(
  value: unknown,
  fallback: number,
  min: number,
  max: number,
): number {
  if (typeof value !== "number" || !Number.isFinite(value)) return fallback;
  return Math.max(min, Math.min(max, Math.floor(value)));
}
