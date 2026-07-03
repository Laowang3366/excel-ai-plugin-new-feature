import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";

import { addOcrExecutors } from "./ocrExecutors";

describe("ocr executors", () => {
  const tempFiles: string[] = [];

  afterEach(() => {
    vi.restoreAllMocks();
    delete process.env.MINERU_API_TOKEN;
    for (const filePath of tempFiles.splice(0)) {
      fs.rmSync(filePath, { force: true });
    }
  });

  it("parses image or pdf files through configured MinerU token first", async () => {
    process.env.MINERU_API_TOKEN = "token";
    const filePath = tempFile("ocr-tool", ".pdf", "pdf");
    const zipBuffer = await mineruZip([
      "# 发票",
      "",
      "| 字段 | 值 |",
      "| --- | --- |",
      "| 发票号码 | 001 |",
    ].join("\n"));

    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      if (url.endsWith("/file-urls/batch")) {
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            batch_id: "batch-1",
            file_urls: ["https://upload.example.com/file"],
          },
        });
      }
      if (url === "https://upload.example.com/file") {
        expect(init?.method).toBe("PUT");
        return new Response("", { status: 200 });
      }
      if (url.endsWith("/extract-results/batch/batch-1")) {
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            extract_result: [{
              file_name: "invoice.pdf",
              state: "done",
              full_zip_url: "https://download.example.com/result.zip",
            }],
          },
        });
      }
      if (url === "https://download.example.com/result.zip") {
        return new Response(zipBuffer, { status: 200 });
      }
      throw new Error(`unexpected url: ${url}`);
    }));

    const result = await executeOcr({ filePaths: [filePath], mode: "invoice" });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({
      provider: "mineru",
      mode: "invoice",
      fileCount: 1,
      text: expect.stringContaining("发票号码"),
      rows: [
        ["字段", "值"],
        ["发票号码", "001"],
      ],
      documents: [{
        filename: "invoice.pdf",
        provider: "mineru",
        text: expect.stringContaining("发票号码"),
      }],
      errors: [],
      fallbacks: [{
        provider: "mineru",
        success: true,
      }],
    });
  });

  it("falls back to free MinerU Agent when token MinerU quota fails", async () => {
    process.env.MINERU_API_TOKEN = "token";
    const filePath = tempFile("ocr-agent", ".pdf", "pdf");

    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      if (url.endsWith("/file-urls/batch")) {
        return jsonResponse({ code: 2001, msg: "额度不足", data: null }, 200);
      }
      if (url.endsWith("/api/v1/agent/parse/file")) {
        expect(init?.method).toBe("POST");
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            task_id: "task-1",
            file_url: "https://upload.example.com/agent-file",
          },
        });
      }
      if (url === "https://upload.example.com/agent-file") {
        expect(init?.method).toBe("PUT");
        return new Response("", { status: 200 });
      }
      if (url.endsWith("/api/v1/agent/parse/task-1")) {
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            task_id: "task-1",
            state: "done",
            markdown_url: "https://download.example.com/agent.md",
          },
        });
      }
      if (url === "https://download.example.com/agent.md") {
        return new Response("# 免费解析\n\n识别到金额 100", { status: 200 });
      }
      throw new Error(`unexpected url: ${url}`);
    }));

    const result = await executeOcr({ filePaths: [filePath] });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({
      provider: "mineru-agent",
      text: expect.stringContaining("识别到金额"),
      warnings: expect.arrayContaining([
        expect.stringContaining("MinerU 标准解析不可用"),
        expect.stringContaining("MinerU 免费 Agent"),
      ]),
      fallbacks: [
        {
          provider: "mineru",
          success: false,
          quotaLikely: true,
        },
        {
          provider: "mineru-agent",
          success: true,
        },
      ],
    });
  });

  it("falls back to local parser when MinerU token and free Agent both fail", async () => {
    process.env.MINERU_API_TOKEN = "token";
    const filePath = tempFile("ocr-local", ".txt", "本地兜底内容\n第二行");

    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/file-urls/batch")) {
        return jsonResponse({ code: 2001, msg: "额度不足", data: null }, 200);
      }
      if (url.endsWith("/api/v1/agent/parse/file")) {
        return jsonResponse({ code: 429, msg: "IP 频率限制", data: null }, 200);
      }
      throw new Error(`unexpected url: ${url}`);
    }));

    const result = await executeOcr({ filePaths: [filePath] });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({
      provider: "local",
      text: expect.stringContaining("本地兜底内容"),
      warnings: expect.arrayContaining([
        expect.stringContaining("MinerU 标准解析不可用"),
        expect.stringContaining("MinerU 免费 Agent 解析不可用"),
        expect.stringContaining("本地免费兜底解析"),
      ]),
      fallbacks: [
        expect.objectContaining({ provider: "mineru", success: false }),
        expect.objectContaining({ provider: "mineru-agent", success: false }),
        expect.objectContaining({ provider: "local", success: true }),
      ],
    });
  });

  it("uses free MinerU Agent when no token is configured", async () => {
    const filePath = tempFile("ocr-no-token", ".pdf", "pdf");

    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/api/v1/agent/parse/file")) {
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            task_id: "task-2",
            file_url: "https://upload.example.com/no-token",
          },
        });
      }
      if (url === "https://upload.example.com/no-token") {
        return new Response("", { status: 200 });
      }
      if (url.endsWith("/api/v1/agent/parse/task-2")) {
        return jsonResponse({
          code: 0,
          msg: "ok",
          data: {
            task_id: "task-2",
            state: "done",
            markdown_url: "https://download.example.com/no-token.md",
          },
        });
      }
      if (url === "https://download.example.com/no-token.md") {
        return new Response("免费链路内容", { status: 200 });
      }
      throw new Error(`unexpected url: ${url}`);
    }));

    const result = await executeOcr({ filePaths: [filePath] });

    expect(result.success).toBe(true);
    expect(result.data).toMatchObject({
      provider: "mineru-agent",
      text: "## " + path.basename(filePath) + "\n免费链路内容",
      fallbacks: [
        {
          provider: "mineru",
          success: false,
          skipped: true,
          reason: expect.stringContaining("Token 未配置"),
        },
        {
          provider: "mineru-agent",
          success: true,
        },
      ],
    });
  });

  function tempFile(prefix: string, ext: string, content: string): string {
    const filePath = path.join(os.tmpdir(), `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2)}${ext}`);
    tempFiles.push(filePath);
    fs.writeFileSync(filePath, content);
    return filePath;
  }
});

async function executeOcr(args: Record<string, unknown>) {
  const executors = new Map();
  addOcrExecutors(executors);
  return await executors.get("ocr.parseDocument")!.execute({
    maxTextChars: 2000,
    maxTableRows: 20,
    ...args,
  });
}

async function mineruZip(markdown: string): Promise<Buffer> {
  const zip = new JSZip();
  zip.file("full.md", markdown);
  return await zip.generateAsync({ type: "nodebuffer" });
}

function jsonResponse(value: unknown, status = 200): Response {
  return new Response(JSON.stringify(value), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}
