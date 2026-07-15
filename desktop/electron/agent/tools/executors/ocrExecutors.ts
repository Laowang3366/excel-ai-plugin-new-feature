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
import { clampNumber } from "../../shared/numberLimits";
import {
  assertRemoteDataProcessingAllowed,
  toRemoteDataPolicyResult,
} from "../../../shared/egressPolicy";
import { parseFilesLocally, type LocalParsedDocument } from "./localDocumentParser";
import {
  buildOcrToolResult,
  type OcrFallbackAttempt,
  type OcrMode,
  type OcrProvider,
  type SelectedOcrDocument,
} from "./ocrExecutorResult";
import { validateArgs } from "./validation";

export interface OcrExecutorDeps {
  getMineruApiToken?: () => string;
  isRemoteDataProcessingEnabled?: () => boolean;
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

export function addOcrExecutors(
  target: Map<string, ToolExecutor>,
  deps: OcrExecutorDeps = {},
): void {
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
      const maxTextChars = clampNumber(args.maxTextChars, {
        fallback: 60_000,
        min: 1_000,
        max: 120_000,
      });
      const maxTableRows = clampNumber(args.maxTableRows, { fallback: 200, min: 0, max: 1_000 });
      const allowTokenMineru = args.allowTokenMineru !== false;
      const allowFreeMineru = args.allowFreeMineru !== false;
      const allowLocalFallback = args.allowLocalFallback !== false;
      const token = getConfiguredMineruToken(deps);
      const remoteDataProcessingEnabled = deps.isRemoteDataProcessingEnabled?.() === true;

      const fallbacks: OcrFallbackAttempt[] = [];
      const warnings: string[] = [];
      const selected: Array<SelectedOcrDocument | undefined> = new Array(filePaths.length);
      let unresolved = filePaths.map((_, index) => index);
      let localDocuments: LocalParsedDocument[] | null = null;

      if (allowLocalFallback) {
        localDocuments = await parseFilesLocally(filePaths);
        fallbacks.push({
          provider: "local",
          success: hasAnyUsefulDocument(localDocuments),
          parsedFiles: localDocuments.filter(hasUsefulDocument).length,
          totalFiles: filePaths.length,
          reason: "已先使用本地解析，仅把本地无法处理的文件交给后续远程服务",
        });
        unresolved = mergeUsefulOcrDocuments(selected, unresolved, localDocuments, "local");
        warnings.push(...localDocuments.flatMap((document) => document.warnings));
      } else {
        fallbacks.push({
          provider: "local",
          success: false,
          skipped: true,
          reason: "调用参数 allowLocalFallback=false，已跳过本地解析",
        });
      }

      if (unresolved.length > 0 && !remoteDataProcessingEnabled) {
        const reason = "远程数据处理已关闭，仅保留本地解析结果";
        fallbacks.push({ provider: "mineru", success: false, skipped: true, reason });
        fallbacks.push({ provider: "mineru-agent", success: false, skipped: true, reason });
        warnings.push(reason);
      }

      if (unresolved.length > 0 && remoteDataProcessingEnabled) {
        try {
          assertRemoteDataProcessingAllowed({
            enabled: true,
            operation: "ocr",
            texts: unresolved.map((index) => localDocuments?.[index]?.text || ""),
          });
        } catch (error) {
          const policyResult = toRemoteDataPolicyResult(error);
          if (policyResult) return policyResult;
          throw error;
        }
      }

      if (unresolved.length > 0 && remoteDataProcessingEnabled && !allowTokenMineru) {
        fallbacks.push({
          provider: "mineru",
          success: false,
          skipped: true,
          reason: "调用参数 allowTokenMineru=false，已跳过配置 token 的 MinerU 标准解析",
        });
      } else if (unresolved.length > 0 && remoteDataProcessingEnabled && !token) {
        fallbacks.push({
          provider: "mineru",
          success: false,
          skipped: true,
          reason: "MinerU API Token 未配置，直接尝试 MinerU 免费 Agent 轻量解析",
        });
      } else if (unresolved.length > 0 && remoteDataProcessingEnabled) {
        const pendingIndices = unresolved;
        const standardAttempt = await tryParseWithProvider("mineru", () =>
          parseFilesWithMineru(
            pendingIndices.map((index) => filePaths[index]),
            token,
          ),
        );
        fallbacks.push(standardAttempt.fallback);
        if (standardAttempt.documents) {
          unresolved = mergeUsefulOcrDocuments(
            selected,
            pendingIndices,
            standardAttempt.documents,
            "mineru",
          );
        }
        if (unresolved.length > 0) {
          warnings.push(
            standardAttempt.fallback.error
              ? `MinerU 标准解析不可用，已进入免费 Agent 补齐: ${standardAttempt.fallback.error}`
              : `MinerU 标准解析仍有 ${unresolved.length} 个文件未完成，已进入免费 Agent 补齐`,
          );
        }
      }

      if (unresolved.length > 0 && remoteDataProcessingEnabled) {
        if (!allowFreeMineru) {
          fallbacks.push({
            provider: "mineru-agent",
            success: false,
            skipped: true,
            reason: "调用参数 allowFreeMineru=false，已跳过 MinerU 免费 Agent 轻量解析",
          });
        } else {
          const pendingIndices = unresolved;
          const agentAttempt = await tryParseWithProvider("mineru-agent", () =>
            parseFilesWithMineruAgent(pendingIndices.map((index) => filePaths[index])),
          );
          fallbacks.push(agentAttempt.fallback);
          if (agentAttempt.documents) {
            unresolved = mergeUsefulOcrDocuments(
              selected,
              pendingIndices,
              agentAttempt.documents,
              "mineru-agent",
            );
          }
          if (unresolved.length > 0) {
            warnings.push(
              agentAttempt.fallback.error
                ? `MinerU 免费 Agent 解析不可用，已进入本地补齐: ${agentAttempt.fallback.error}`
                : `MinerU 免费 Agent 仍有 ${unresolved.length} 个文件未完成，已进入本地补齐`,
            );
          }
        }
      }

      if (unresolved.length > 0) {
        if (!allowLocalFallback || !localDocuments) {
          return {
            success: false,
            error: `仍有 ${unresolved.length} 个文件未解析，且本地解析不可用`,
            data: { fallbacks },
          };
        }
        for (const index of unresolved) {
          const document = localDocuments[index];
          if (document) selected[index] = { document, provider: "local" };
        }
      }

      return buildOcrToolResult({
        fallbacks,
        filePaths,
        maxTableRows,
        maxTextChars,
        mode,
        selected,
        warnings,
      });
    },
  });
}

async function tryParseWithProvider(
  provider: OcrProvider,
  parse: () => Promise<MineruParsedDocument[]>,
): Promise<{ fallback: OcrFallbackAttempt; documents?: MineruParsedDocument[] }> {
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

function mergeUsefulOcrDocuments(
  selected: Array<
    { document: MineruParsedDocument | LocalParsedDocument; provider: OcrProvider } | undefined
  >,
  targetIndices: number[],
  documents: Array<MineruParsedDocument | LocalParsedDocument>,
  provider: OcrProvider,
): number[] {
  const unresolved: number[] = [];
  for (let index = 0; index < targetIndices.length; index++) {
    const document = documents[index];
    if (document && hasUsefulDocument(document)) {
      selected[targetIndices[index]] = { document, provider };
    } else {
      unresolved.push(targetIndices[index]);
    }
  }
  return unresolved;
}

function isQuotaLikeError(error: string): boolean {
  return /quota|limit|limited|rate|credit|balance|insufficient|次数|额度|余额|限额|频率|too many|429/i.test(
    error,
  );
}
