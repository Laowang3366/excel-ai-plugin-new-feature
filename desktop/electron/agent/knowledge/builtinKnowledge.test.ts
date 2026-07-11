import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { createHash } from "crypto";
import { afterEach, describe, expect, it, vi } from "vitest";

import { indexBuiltinKnowledge } from "./builtinKnowledge";

let tempRoots: string[] = [];

afterEach(() => {
  vi.restoreAllMocks();
  for (const root of tempRoots) {
    fs.rmSync(root, { recursive: true, force: true });
  }
  tempRoots = [];
});

describe("indexBuiltinKnowledge", () => {
  it("bundles only the synthesized methodology with a valid content hash", () => {
    const knowledgeRoot = path.resolve(__dirname, "../../../public/knowledge");
    const manifest = JSON.parse(
      fs.readFileSync(path.join(knowledgeRoot, "builtin-knowledge.json"), "utf-8"),
    ) as { version: number; files: Array<{ path: string; sha256: string }> };

    expect(manifest.version).toBe(2);
    expect(manifest.files).toHaveLength(1);
    expect(manifest.files[0].path).toBe("excel-wps-formula-problem-solving-methodology.md");
    const content = fs
      .readFileSync(path.join(knowledgeRoot, manifest.files[0].path), "utf-8")
      .replace(/\r\n/g, "\n");
    expect(content).toContain("先生成“索引、掩码或状态”，最后取值");
    expect(content).toContain(
      "LET 定义输入 -> 规范化 -> 构造索引/掩码/状态 -> 执行局部变换 -> 展平或累积 -> 重塑 -> HSTACK/VSTACK 输出",
    );
    expect(content).toContain("追求“最小充分公式”，不靠普通函数堆叠");
    expect(content).toContain("正则不只是展示文本的清洗工具");
    expect(content).toContain("结果正确、结构精简、无重复计算、动态可扩展和可维护，缺一不可");
    expect(createHash("sha256").update(content).digest("hex")).toBe(manifest.files[0].sha256);
    expect(fs.readdirSync(knowledgeRoot).sort()).toEqual([
      "builtin-knowledge.json",
      "excel-wps-formula-problem-solving-methodology.md",
    ]);
  });

  it("indexes the methodology before removing unlisted builtin sources", async () => {
    const root = createBuiltinRoot();
    const knowledgeRoot = path.join(root, "public", "knowledge");
    const methodologyPath = path.join(
      knowledgeRoot,
      "excel-wps-formula-problem-solving-methodology.md",
    );
    const calls: string[] = [];
    const indexer = {
      indexFile: vi.fn(async (sourcePath: string) => {
        calls.push(`index:${path.basename(sourcePath)}`);
        return {
          sourcePath,
          success: true,
          entryCount: 1,
          durationMs: 1,
        };
      }),
      deleteSource: vi.fn(async (sourcePath: string) => {
        calls.push(`delete:${path.basename(sourcePath)}`);
      }),
      listSources: vi.fn(() => [
        {
          sourcePath: "C:\\old-workspace\\desktop\\public\\knowledge\\stale.md",
          sourceName: "stale.md",
        },
        {
          sourcePath: methodologyPath,
          sourceName: "excel-wps-formula-problem-solving-methodology.md",
        },
        {
          sourcePath: "D:\\user-documents\\notes.md",
          sourceName: "notes.md",
        },
      ]),
    };

    vi.spyOn(process, "cwd").mockReturnValue(root);
    const results = await indexBuiltinKnowledge(indexer as any);

    expect(results).toHaveLength(1);
    expect(indexer.indexFile).toHaveBeenCalledWith(methodologyPath, {
      skipUnchanged: true,
      knownFileHash: "methodology-hash",
    });
    expect(indexer.deleteSource).toHaveBeenCalledOnce();
    expect(indexer.deleteSource).toHaveBeenCalledWith(
      "C:\\old-workspace\\desktop\\public\\knowledge\\stale.md",
    );
    expect(calls[0]).toBe("index:excel-wps-formula-problem-solving-methodology.md");
  });

  it("keeps existing sources when the replacement cannot be indexed", async () => {
    const root = createBuiltinRoot();
    const indexer = {
      indexFile: vi.fn(async (sourcePath: string) => ({
        sourcePath,
        success: false,
        error: "embedding unavailable",
        entryCount: 0,
        durationMs: 1,
      })),
      deleteSource: vi.fn(),
      listSources: vi.fn(),
    };

    vi.spyOn(process, "cwd").mockReturnValue(root);
    await indexBuiltinKnowledge(indexer as any);

    expect(indexer.deleteSource).not.toHaveBeenCalled();
    expect(indexer.listSources).not.toHaveBeenCalled();
  });
});

function createBuiltinRoot(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "builtin-knowledge-"));
  tempRoots.push(root);
  const knowledgeRoot = path.join(root, "public", "knowledge");
  fs.mkdirSync(knowledgeRoot, { recursive: true });
  fs.writeFileSync(
    path.join(knowledgeRoot, "builtin-knowledge.json"),
    JSON.stringify({
      version: 2,
      files: [
        {
          path: "excel-wps-formula-problem-solving-methodology.md",
          sha256: "methodology-hash",
        },
      ],
    }),
  );
  fs.writeFileSync(
    path.join(knowledgeRoot, "excel-wps-formula-problem-solving-methodology.md"),
    "# Formula methodology",
  );
  return root;
}
