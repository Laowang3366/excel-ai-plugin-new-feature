import * as fs from "fs";
import * as path from "path";
import { createHash } from "crypto";
import { strFromU8, unzipSync } from "fflate";
import { extractMarkdownTables } from "../shared/markdownTables";

export interface MineruParsedDocument {
  filename: string;
  text: string;
  rows: string[][];
  fullZipUrl?: string;
  error?: string;
}

interface MineruBatchResultItem {
  file_name?: string;
  state?: string;
  full_zip_url?: string;
  err_msg?: string;
}

const MINERU_API_BASE = "https://mineru.net/api/v4";
export const DEFAULT_POLL_TIMEOUT_MS = 180_000;
export const DEFAULT_POLL_INTERVAL_MS = 3_000;
export const DEFAULT_REQUEST_TIMEOUT_MS = 30_000;

export async function parseFilesWithMineru(
  filePaths: string[],
  token: string,
  options?: { timeoutMs?: number; pollIntervalMs?: number },
): Promise<MineruParsedDocument[]> {
  const normalizedToken = token.trim();
  if (!normalizedToken) {
    throw new Error("MinerU API Token 未配置");
  }
  if (filePaths.length === 0) {
    throw new Error("请先选择要识别的文件");
  }

  const apply = await requestBatchUploadUrls(filePaths, normalizedToken);
  await uploadFilesToSignedUrls(filePaths, apply.fileUrls);
  const results = await waitForBatchResults(
    apply.batchId,
    normalizedToken,
    options?.timeoutMs ?? DEFAULT_POLL_TIMEOUT_MS,
    options?.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS,
  );

  return await Promise.all(
    results.map(async (item, index) => {
      const filename = item.file_name || path.basename(filePaths[index] || "");
      if (item.state === "failed") {
        return {
          filename,
          text: "",
          rows: [],
          error: item.err_msg || "MinerU 解析失败",
        };
      }
      if (!item.full_zip_url) {
        return {
          filename,
          text: "",
          rows: [],
          error: `MinerU 未返回解析结果压缩包，状态: ${item.state || "unknown"}`,
        };
      }
      try {
        const markdown = await downloadFullMarkdown(item.full_zip_url);
        return {
          filename,
          text: markdown,
          rows: extractMarkdownTables(markdown),
          fullZipUrl: item.full_zip_url,
        };
      } catch (error: any) {
        return {
          filename,
          text: "",
          rows: [],
          fullZipUrl: item.full_zip_url,
          error: error?.message || "读取 MinerU 解析结果失败",
        };
      }
    }),
  );
}

async function requestBatchUploadUrls(
  filePaths: string[],
  token: string,
): Promise<{ batchId: string; fileUrls: string[] }> {
  const response = await fetchWithTimeout(`${MINERU_API_BASE}/file-urls/batch`, {
    method: "POST",
    headers: mineruJsonHeaders(token),
    body: JSON.stringify({
      files: filePaths.map((filePath, index) => ({
        name: path.basename(filePath),
        data_id: buildDataId(filePath, index),
      })),
      model_version: "vlm",
    }),
  });
  const json = await readMineruJson(response);
  const batchId = json.data?.batch_id;
  const fileUrls = json.data?.file_urls;
  if (typeof batchId !== "string" || !Array.isArray(fileUrls)) {
    throw new Error(`MinerU 上传签名响应异常: ${JSON.stringify(json).slice(0, 300)}`);
  }
  if (fileUrls.length !== filePaths.length) {
    throw new Error("MinerU 返回的上传链接数量与文件数量不一致");
  }
  return { batchId, fileUrls: fileUrls.map(String) };
}

async function uploadFilesToSignedUrls(filePaths: string[], fileUrls: string[]): Promise<void> {
  for (let i = 0; i < filePaths.length; i++) {
    await uploadFileToSignedUrl(filePaths[i], fileUrls[i]);
  }
}

export async function uploadFileToSignedUrl(filePath: string, fileUrl: string): Promise<void> {
  const fileBuffer = await fs.promises.readFile(filePath);
  const response = await fetchWithTimeout(fileUrl, {
    method: "PUT",
    body: fileBuffer,
  });
  if (!response.ok) {
    throw new Error(
      `MinerU 文件上传失败 (${response.status}): ${await safeResponseText(response)}`,
    );
  }
}

async function waitForBatchResults(
  batchId: string,
  token: string,
  timeoutMs: number,
  pollIntervalMs: number,
): Promise<MineruBatchResultItem[]> {
  const startedAt = Date.now();
  while (Date.now() - startedAt <= timeoutMs) {
    const remainingMs = Math.max(1, timeoutMs - (Date.now() - startedAt));
    const response = await fetchWithTimeout(
      `${MINERU_API_BASE}/extract-results/batch/${encodeURIComponent(batchId)}`,
      {
        method: "GET",
        headers: mineruJsonHeaders(token),
      },
      Math.min(DEFAULT_REQUEST_TIMEOUT_MS, remainingMs),
    );
    const json = await readMineruJson(response);
    const items = readBatchResultItems(json);
    if (!Array.isArray(items)) {
      throw new Error(`MinerU 批量结果响应异常: ${JSON.stringify(json).slice(0, 300)}`);
    }

    const states = items.map((item: MineruBatchResultItem) => item.state || "");
    const finished = states.every((state) => state === "done" || state === "failed");
    if (finished) return items;
    await delay(pollIntervalMs);
  }
  throw new Error("MinerU 解析超时");
}

function readBatchResultItems(json: any): MineruBatchResultItem[] | null {
  const data = json?.data;
  const candidates = [
    data?.extract_result,
    data?.extract_results,
    data?.results,
    data?.files,
    data,
  ];
  for (const candidate of candidates) {
    if (Array.isArray(candidate)) return candidate;
  }
  return null;
}

async function downloadFullMarkdown(zipUrl: string): Promise<string> {
  const response = await fetchWithTimeout(zipUrl);
  if (!response.ok) {
    throw new Error(
      `下载 MinerU 结果压缩包失败 (${response.status}): ${await safeResponseText(response)}`,
    );
  }
  const files = unzipSync(new Uint8Array(await response.arrayBuffer()));
  const markdown = findFullMarkdownFile(files);
  if (!markdown) {
    throw new Error("MinerU 结果压缩包中未找到 full.md");
  }
  return strFromU8(markdown);
}

export async function downloadMarkdownText(markdownUrl: string): Promise<string> {
  const response = await fetchWithTimeout(markdownUrl);
  if (!response.ok) {
    throw new Error(
      `下载 MinerU Agent Markdown 失败 (${response.status}): ${await safeResponseText(response)}`,
    );
  }
  return await response.text();
}

function findFullMarkdownFile(files: Record<string, Uint8Array>): Uint8Array | null {
  const entries = Object.entries(files);
  return (
    entries.find(([name]) => /(^|\/)full\.md$/i.test(name))?.[1] ??
    entries.find(([name]) => /\.md$/i.test(name))?.[1] ??
    null
  );
}

function mineruJsonHeaders(token: string): Record<string, string> {
  return {
    "Content-Type": "application/json",
    Accept: "*/*",
    Authorization: `Bearer ${token}`,
  };
}

export async function readMineruJson(response: Response, provider = "MinerU"): Promise<any> {
  const text = await response.text();
  let json: any;
  try {
    json = JSON.parse(text);
  } catch {
    throw new Error(`${provider} 返回非 JSON: ${formatErrorText(text)}`);
  }
  if (!response.ok || json.code !== 0) {
    throw new Error(`${provider} 请求失败: ${json.msg || response.statusText || response.status}`);
  }
  return json;
}

function buildDataId(filePath: string, index: number): string {
  const hash = createHash("sha256")
    .update(`${filePath}:${Date.now()}:${index}`, "utf8")
    .digest("hex")
    .slice(0, 16);
  return `ocr_${hash}`;
}

export function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export async function fetchWithTimeout(
  url: string,
  init?: RequestInit,
  timeoutMs = DEFAULT_REQUEST_TIMEOUT_MS,
): Promise<Response> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), Math.max(1, timeoutMs));
  try {
    return await fetch(url, { ...init, signal: controller.signal });
  } catch (error) {
    if (controller.signal.aborted) {
      throw new Error(`MinerU 网络请求超时 (${timeoutMs}ms)`);
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function safeResponseText(response: Response): Promise<string> {
  try {
    return formatErrorText(await response.text());
  } catch {
    return "";
  }
}

function formatErrorText(value: string): string {
  return value
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 300);
}

/** Thin wrapper keeps public export on this module without circular import. */
export async function parseFilesWithMineruAgent(
  filePaths: string[],
  options?: { timeoutMs?: number; pollIntervalMs?: number },
): Promise<MineruParsedDocument[]> {
  const { parseFilesWithMineruAgent: runAgent } = await import("./mineruOcrAgent");
  return runAgent(filePaths, options);
}
