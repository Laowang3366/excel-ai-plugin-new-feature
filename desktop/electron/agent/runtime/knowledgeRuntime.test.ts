import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { afterEach, describe, expect, it } from "vitest";

import { getKnowledgeRetriever, resetKnowledgeRegistry } from "../knowledge";
import type { AIClientConfig } from "../providers/aiClient";
import {
  initializeKnowledgeRuntime,
  resetKnowledgeRuntime,
} from "./knowledgeRuntime";

describe("knowledgeRuntime", () => {
  const tempDirs: string[] = [];

  afterEach(() => {
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
});
