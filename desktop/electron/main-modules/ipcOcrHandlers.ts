import { trustedIpcMain as ipcMain } from "../shared/trustedIpc";
import { parseFilesLocally } from "../agent/tools/executors/localDocumentParser";
import { validateInput, OcrRecognizeInput } from "../shared/ipcSchemas";
import {
  assertRemoteDataProcessingAllowed,
  isRemoteDataProcessingEnabled,
  type RemoteDataTransferSummary,
} from "../shared/egressPolicy";
import { assertAuthorizedPath, createPathAuthorizer } from "./ipcPathSecurity";
import { getRuntimeSettingValue } from "./settingsManager";
import {
  parseFilesWithMineru,
  parseFilesWithMineruAgent,
  type MineruParsedDocument,
} from "./mineruOcr";
import {
  normalizeOcrMode,
  isLikelyInvoiceFileList,
  type OcrVisionResult,
} from "./ocrModeDetection";
import { guardDataOperation } from "./dataMaintenance";
import {
  buildOcrResultFromDocuments,
  emptyOcrResult,
  formatOcrDocumentErrors,
} from "./ocrDocumentResultBuilder";

export { normalizeOcrVisionResult } from "./ocrDocumentResultBuilder";

export function registerOcrIpcHandler(
  pathAuthorizer: ReturnType<typeof createPathAuthorizer>,
  isDataMaintenanceInProgress?: () => boolean,
): void {
  try {
    ipcMain.removeHandler("ocr:recognize");
  } catch {
    // Handler may not exist on first registration.
  }

  ipcMain.handle(
    "ocr:recognize",
    guardDataOperation(
      isDataMaintenanceInProgress,
      async (_event, mode: unknown, filePaths: unknown) => {
        try {
          const validated = validateInput(OcrRecognizeInput, { mode, filePaths });
          const authorizedFilePaths = validated.filePaths.map((filePath) =>
            assertAuthorizedPath(pathAuthorizer, filePath),
          );
          return await recognizeWithOcrFallbacks(validated.mode, authorizedFilePaths);
        } catch (err: any) {
          return emptyOcrResult(normalizeOcrMode(mode), [err?.message || "OCR 识别失败"]);
        }
      },
    ),
  );
}

export async function recognizeWithOcrFallbacks(
  rawMode: unknown,
  rawFilePaths: unknown,
): Promise<OcrVisionResult> {
  const mode = normalizeOcrMode(rawMode);
  const filePaths = normalizeOcrFilePaths(rawFilePaths);
  const effectiveMode =
    mode === "invoice" || isLikelyInvoiceFileList(filePaths) ? "invoice" : "image";

  const remoteEnabled = isRemoteDataProcessingEnabled(
    getRuntimeSettingValue("remoteDataProcessingEnabled"),
  );
  const parsed = await parseFilesWithOcrFallbacks(filePaths, remoteEnabled);
  if (!hasAnyUsefulParsedDocument(parsed.documents)) {
    return emptyOcrResult(effectiveMode, [
      "未提取到可用 OCR 文本或表格，无法抽取字段",
      ...parsed.errors,
      ...formatParsedDocumentErrors(parsed.documents),
    ]);
  }

  const result = await buildOcrResultFromDocuments(parsed.documents, effectiveMode, remoteEnabled);

  return {
    ...result,
    remoteProcessing: [...parsed.remoteProcessing, ...(result.remoteProcessing || [])],
    errors: [...result.errors, ...parsed.errors, ...formatParsedDocumentErrors(parsed.documents)],
  };
}

async function parseFilesWithOcrFallbacks(
  filePaths: string[],
  remoteEnabled: boolean,
): Promise<{
  documents: MineruParsedDocument[];
  errors: string[];
  remoteProcessing: RemoteDataTransferSummary[];
}> {
  const errors: string[] = [];
  const remoteProcessing: RemoteDataTransferSummary[] = [];
  const mineruToken = getConfiguredMineruToken();
  const selected: Array<MineruParsedDocument | undefined> = new Array(filePaths.length);
  const localDocuments = await parseFilesLocally(filePaths);
  let unresolved = mergeUsefulDocuments(
    selected,
    filePaths.map((_, index) => index),
    localDocuments,
  );

  if (unresolved.length > 0 && !remoteEnabled) {
    errors.push("远程数据处理已关闭，本次 OCR 仅使用本地解析");
  }

  if (unresolved.length > 0 && remoteEnabled) {
    try {
      assertRemoteDataProcessingAllowed({
        enabled: true,
        operation: "ocr",
        texts: unresolved.map((index) => localDocuments[index]?.text || ""),
      });
    } catch (error) {
      errors.push(error instanceof Error ? error.message : String(error));
      unresolved = [];
    }
  }

  if (unresolved.length > 0 && remoteEnabled && mineruToken) {
    const pendingCount = unresolved.length;
    try {
      const pendingIndices = unresolved;
      const documents = await parseFilesWithMineru(
        pendingIndices.map((index) => filePaths[index]),
        mineruToken,
      );
      unresolved = mergeUsefulDocuments(selected, pendingIndices, documents);
      const completedCount = pendingCount - unresolved.length;
      if (completedCount > 0) {
        remoteProcessing.push({
          operation: "ocr",
          service: "MinerU",
          destination: "mineru.net",
          dataSummary: `${completedCount} 个文件`,
        });
      }
      if (unresolved.length > 0)
        errors.push(formatOcrDocumentErrors(documents) || "MinerU 标准解析存在未完成文件");
    } catch (error: any) {
      errors.push(`MinerU 标准解析失败：${error?.message || "未知错误"}`);
    }
  }

  if (unresolved.length > 0 && remoteEnabled) {
    const pendingIndices = unresolved;
    const pendingCount = unresolved.length;
    try {
      const documents = await parseFilesWithMineruAgent(
        pendingIndices.map((index) => filePaths[index]),
      );
      unresolved = mergeUsefulDocuments(selected, pendingIndices, documents);
      const completedCount = pendingCount - unresolved.length;
      if (completedCount > 0) {
        remoteProcessing.push({
          operation: "ocr",
          service: "MinerU Agent",
          destination: "mineru.net",
          dataSummary: `${completedCount} 个文件`,
        });
      }
      if (unresolved.length > 0)
        errors.push(formatOcrDocumentErrors(documents) || "MinerU 免费解析存在未完成文件");
    } catch (error: any) {
      errors.push(`MinerU 免费解析失败：${error?.message || "未知错误"}`);
    }
  }

  if (unresolved.length > 0) {
    for (const index of unresolved) {
      selected[index] = localDocuments[index];
    }
  }

  const documents = selected.filter((document): document is MineruParsedDocument =>
    Boolean(document),
  );
  return {
    documents,
    errors,
    remoteProcessing,
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

function formatParsedDocumentErrors(
  documents: Array<{ filename: string; error?: string }>,
): string[] {
  return documents
    .filter(
      (document) =>
        document.error &&
        !/^local_ocr_unsupported|local_unsupported|local_empty$/.test(document.error),
    )
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

function getConfiguredMineruToken(): string {
  const configured =
    getRuntimeSettingValue("mineruApiToken") || getRuntimeSettingValue("ocrMineruApiToken");
  const tokenFromSettings = typeof configured === "string" ? configured.trim() : "";
  return tokenFromSettings || (process.env.MINERU_API_TOKEN || "").trim();
}
