import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getKnowledgeRetriever, getKnowledgeStore, resetKnowledgeRegistry } from "../knowledge";
import type { AIClientConfig } from "../providers/aiClient";
import {
  initializeKnowledgeRuntime,
  reloadKnowledgeRuntime,
  resetKnowledgeRuntime,
} from "./knowledgeRuntime";

describe("knowledgeRuntime", () => {
  const tempDirs: string[] = [];
  const originalFetch = globalThis.fetch;

  beforeEach(() => {
    mockEmbeddingFetch();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
    vi.restoreAllMocks();
    resetKnowledgeRuntime();
    resetKnowledgeRegistry();
    for (const dir of tempDirs.splice(0)) {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });

  it("rebuilds knowledge services when the embedding provider config changes", async () => {
    const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-runtime-"));
    tempDirs.push(dataRoot);

    const openaiConfig: AIClientConfig = {
      provider: "openai",
      apiKey: "sk-old",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };
    const qwenConfig: AIClientConfig = {
      provider: "qwen",
      apiKey: "sk-new",
      baseUrl: "https://dashscope.example.test/compatible-mode/v1",
      model: "qwen-plus",
    };

    const first = await initializeKnowledgeRuntime(openaiConfig, dataRoot);
    expect(first.embedder?.getProfile()).toMatchObject({
      provider: "openai",
      model: "text-embedding-3-small",
    });
    const firstRetriever = getKnowledgeRetriever();
    expect(firstRetriever).toBe(first.retriever);

    const second = await initializeKnowledgeRuntime(qwenConfig, dataRoot);
    expect(second.embedder?.getProfile()).toMatchObject({
      provider: "qwen",
      model: "text-embedding-v2",
    });
    expect(second.embedder).not.toBe(first.embedder);
    expect(getKnowledgeRetriever()).toBe(second.retriever);
    expect(getKnowledgeRetriever()).not.toBe(firstRetriever);
  });

  it("indexes builtin knowledge on startup", async () => {
    const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-runtime-"));
    tempDirs.push(dataRoot);

    const config: AIClientConfig = {
      provider: "openai",
      apiKey: "sk-test",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };

    const runtime = await initializeKnowledgeRuntime(config, dataRoot);
    const sources = runtime.store?.listSources() ?? [];
    const builtin = sources.find((source) =>
      source.sourceName === "excel-wps-formula-problem-solving-methodology.md"
    );

    expect(builtin).toBeTruthy();
    expect(builtin?.entryCount).toBeGreaterThan(0);
    const entries = runtime.store?.getEntriesBySource(builtin!.sourcePath) ?? [];
    expect(entries.length).toBeGreaterThan(0);
    expect(entries.every((entry) => entry.embedding)).toBe(true);
    expect(entries.every((entry) => entry.embeddingProvider === "openai")).toBe(true);
    expect(entries.every((entry) => entry.embeddingModel === "text-embedding-3-small")).toBe(true);
    expect(entries.some((entry) => entry.content.includes("通用变换链"))).toBe(true);
    expect(entries.some((entry) => entry.content.includes("不能直接视为当前任务答案"))).toBe(true);
    expect(entries.some((entry) => entry.content.includes("控制数组与广播"))).toBe(true);
    expect(entries.some((entry) => entry.content.includes("组内密集排名"))).toBe(true);
    expect(entries.some((entry) => entry.content.includes("大数组分块处理"))).toBe(true);
  });

  it("reuses unchanged builtin knowledge embeddings across startups", async () => {
    const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-runtime-"));
    tempDirs.push(dataRoot);

    const config: AIClientConfig = {
      provider: "openai",
      apiKey: "sk-test",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };

    await initializeKnowledgeRuntime(config, dataRoot);
    resetKnowledgeRuntime();

    const fetchCallsAfterFirstStartup = vi.mocked(globalThis.fetch).mock.calls.length;
    expect(fetchCallsAfterFirstStartup).toBeGreaterThan(0);

    await initializeKnowledgeRuntime(config, dataRoot);

    expect(vi.mocked(globalThis.fetch).mock.calls.length).toBe(fetchCallsAfterFirstStartup);
  });

  it("keeps the active runtime when a replacement cannot initialize", async () => {
    const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), "knowledge-runtime-"));
    tempDirs.push(dataRoot);
    const invalidRoot = path.join(dataRoot, "not-a-directory");
    fs.writeFileSync(invalidRoot, "file");
    const config: AIClientConfig = {
      provider: "openai",
      apiKey: "sk-test",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };

    const initial = await initializeKnowledgeRuntime(config, path.join(dataRoot, "valid"));
    const activeStore = initial.store;
    expect(activeStore).not.toBeNull();

    const reloaded = await reloadKnowledgeRuntime(config, invalidRoot);

    expect(reloaded.store).toBe(activeStore);
    expect(getKnowledgeStore()).toBe(activeStore);
    expect(reloaded.error).toBeTruthy();
    expect(activeStore?.isInitialized()).toBe(true);
  });
});

function mockEmbeddingFetch(): void {
  globalThis.fetch = vi.fn(async (_url, init) => {
    const body = JSON.parse(String(init?.body || "{}")) as { input?: string[] | string };
    const inputCount = Array.isArray(body.input) ? body.input.length : 1;
    const embedding = Array.from({ length: 1536 }, (_, index) => index / 1536);
    return {
      ok: true,
      json: async () => ({
        data: Array.from({ length: inputCount }, () => ({ embedding })),
      }),
    } as Response;
  }) as typeof fetch;
}
