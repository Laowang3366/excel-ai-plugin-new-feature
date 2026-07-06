import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import JSZip from "jszip";
import { afterEach, describe, expect, it, vi } from "vitest";

import { extractMarkdownTables, parseFilesWithMineru, parseFilesWithMineruAgent } from "./mineruOcr";

describe("mineruOcr", () => {
  const tempFiles: string[] = [];

  afterEach(() => {
    vi.restoreAllMocks();
    for (const filePath of tempFiles.splice(0)) {
      fs.rmSync(filePath, { force: true });
    }
  });

  it("extracts markdown table rows", () => {
    const markdown = [
      "发票明细",
      "",
      "| 项目 | 金额 |",
      "| --- | ---: |",
      "| 服务费 | 100.00 |",
      "| 税额 | 6.00 |",
    ].join("\n");

    expect(extractMarkdownTables(markdown)).toEqual([
      ["项目", "金额"],
      ["服务费", "100.00"],
      ["税额", "6.00"],
    ]);
  });

  it("parses files through MinerU signed-url flow", async () => {
    const filePath = path.join(os.tmpdir(), `mineru-test-${Date.now()}.pdf`);
    tempFiles.push(filePath);
    fs.writeFileSync(filePath, "pdf");

    const zip = new JSZip();
    zip.file("sample/full.md", [
      "# 发票",
      "",
      "| 字段 | 值 |",
      "| --- | --- |",
      "| 发票号码 | 001 |",
    ].join("\n"));
    const zipBuffer = await zip.generateAsync({ type: "nodebuffer" });

    // @MOCK_INTERFACE: simulates MinerU signed-url upload, polling, and zip download HTTP endpoints.
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url.endsWith("/file-urls/batch")) {
        expect(init?.method).toBe("POST");
        expect(JSON.parse(String(init?.body))).toMatchObject({
          files: [{ name: path.basename(filePath) }],
          model_version: "vlm",
        });
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
    });
    vi.stubGlobal("fetch", fetchMock);

    const result = await parseFilesWithMineru([filePath], "token", {
      timeoutMs: 1000,
      pollIntervalMs: 1,
    });

    expect(result).toEqual([{
      filename: "invoice.pdf",
      text: expect.stringContaining("发票号码"),
      rows: [
        ["字段", "值"],
        ["发票号码", "001"],
      ],
      fullZipUrl: "https://download.example.com/result.zip",
    }]);
  });

  it("parses files through MinerU Agent lightweight flow", async () => {
    const filePath = path.join(os.tmpdir(), `mineru-agent-${Date.now()}.pdf`);
    tempFiles.push(filePath);
    fs.writeFileSync(filePath, "pdf");

    // @MOCK_INTERFACE: simulates MinerU Agent lightweight parse, upload, poll, and markdown download endpoints.
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url.endsWith("/api/v1/agent/parse/file")) {
        expect(init?.method).toBe("POST");
        expect(JSON.parse(String(init?.body))).toMatchObject({
          file_name: path.basename(filePath),
          language: "ch",
          enable_table: true,
        });
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
        return new Response([
          "# 轻量解析",
          "",
          "| 字段 | 值 |",
          "| --- | --- |",
          "| 金额 | 100 |",
        ].join("\n"), { status: 200 });
      }
      throw new Error(`unexpected url: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock);

    const result = await parseFilesWithMineruAgent([filePath], {
      timeoutMs: 1000,
      pollIntervalMs: 1,
    });

    expect(result).toEqual([{
      filename: path.basename(filePath),
      text: expect.stringContaining("金额"),
      rows: [
        ["字段", "值"],
        ["金额", "100"],
      ],
    }]);
  });
});

function jsonResponse(value: unknown): Response {
  return new Response(JSON.stringify(value), {
    status: 200,
    headers: { "Content-Type": "application/json" },
  });
}
