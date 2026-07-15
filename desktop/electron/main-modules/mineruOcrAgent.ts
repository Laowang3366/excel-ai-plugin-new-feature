import * as path from "path";
import { extractMarkdownTables } from "../shared/markdownTables";
import type { MineruParsedDocument } from "./mineruOcr";
import {
  DEFAULT_POLL_INTERVAL_MS,
  DEFAULT_POLL_TIMEOUT_MS,
  DEFAULT_REQUEST_TIMEOUT_MS,
  delay,
  downloadMarkdownText,
  fetchWithTimeout,
  readMineruJson,
  uploadFileToSignedUrl,
} from "./mineruOcr";

const MINERU_AGENT_API_BASE = "https://mineru.net/api/v1/agent";

interface MineruAgentTaskResult {
  task_id?: string;
  state?: string;
  markdown_url?: string;
  err_msg?: string;
  err_code?: number;
}

export async function parseFilesWithMineruAgent(
  filePaths: string[],
  options?: { timeoutMs?: number; pollIntervalMs?: number },
): Promise<MineruParsedDocument[]> {
  if (filePaths.length === 0) {
    throw new Error("请先选择要识别的文件");
  }

  const documents: MineruParsedDocument[] = [];
  for (const filePath of filePaths) {
    documents.push(
      await parseFileWithMineruAgent(filePath, {
        timeoutMs: options?.timeoutMs ?? DEFAULT_POLL_TIMEOUT_MS,
        pollIntervalMs: options?.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS,
      }),
    );
  }
  return documents;
}

async function parseFileWithMineruAgent(
  filePath: string,
  options: { timeoutMs: number; pollIntervalMs: number },
): Promise<MineruParsedDocument> {
  const filename = path.basename(filePath);
  const apply = await requestAgentUploadUrl(filePath);
  await uploadFileToSignedUrl(filePath, apply.fileUrl);
  const result = await waitForAgentResult(apply.taskId, options.timeoutMs, options.pollIntervalMs);

  if (result.state === "failed") {
    return {
      filename,
      text: "",
      rows: [],
      error: formatAgentError(result),
    };
  }
  if (!result.markdown_url) {
    return {
      filename,
      text: "",
      rows: [],
      error: `MinerU Agent 未返回 Markdown 链接，状态: ${result.state || "unknown"}`,
    };
  }

  try {
    const markdown = await downloadMarkdownText(result.markdown_url);
    return {
      filename,
      text: markdown,
      rows: extractMarkdownTables(markdown),
    };
  } catch (error: any) {
    return {
      filename,
      text: "",
      rows: [],
      error: error?.message || "读取 MinerU Agent 解析结果失败",
    };
  }
}

async function requestAgentUploadUrl(
  filePath: string,
): Promise<{ taskId: string; fileUrl: string }> {
  const response = await fetchWithTimeout(`${MINERU_AGENT_API_BASE}/parse/file`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "*/*",
    },
    body: JSON.stringify({
      file_name: path.basename(filePath),
      language: "ch",
      enable_table: true,
      is_ocr: true,
      enable_formula: true,
    }),
  });
  const json = await readMineruJson(response, "MinerU Agent");
  const taskId = json.data?.task_id;
  const fileUrl = json.data?.file_url;
  if (typeof taskId !== "string" || typeof fileUrl !== "string") {
    throw new Error(`MinerU Agent 上传签名响应异常: ${JSON.stringify(json).slice(0, 300)}`);
  }
  return { taskId, fileUrl };
}

async function waitForAgentResult(
  taskId: string,
  timeoutMs: number,
  pollIntervalMs: number,
): Promise<MineruAgentTaskResult> {
  const startedAt = Date.now();
  while (Date.now() - startedAt <= timeoutMs) {
    const remainingMs = Math.max(1, timeoutMs - (Date.now() - startedAt));
    const response = await fetchWithTimeout(
      `${MINERU_AGENT_API_BASE}/parse/${encodeURIComponent(taskId)}`,
      {
        method: "GET",
        headers: { Accept: "*/*" },
      },
      Math.min(DEFAULT_REQUEST_TIMEOUT_MS, remainingMs),
    );
    const json = await readMineruJson(response, "MinerU Agent");
    const data = json.data as MineruAgentTaskResult | undefined;
    if (!data || typeof data !== "object") {
      throw new Error(`MinerU Agent 查询响应异常: ${JSON.stringify(json).slice(0, 300)}`);
    }
    if (data.state === "done" || data.state === "failed") return data;
    await delay(pollIntervalMs);
  }
  throw new Error("MinerU Agent 解析超时");
}

function formatAgentError(result: MineruAgentTaskResult): string {
  const parts = [
    result.err_msg,
    typeof result.err_code === "number" ? `err_code=${result.err_code}` : "",
  ].filter(Boolean);
  return parts.length > 0 ? parts.join(" ") : "MinerU Agent 解析失败";
}
