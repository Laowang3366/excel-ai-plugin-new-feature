import * as path from "path";
import * as fs from "fs";
import type { AIClientConfig } from "../providers/aiClient";
import {
  EmbeddingService,
  KnowledgeIndexer,
  KnowledgeWriter,
  Retriever,
  SqliteStore,
  resetKnowledgeRegistry,
  setKnowledgeIndexer,
  setKnowledgeRetriever,
  setKnowledgeStore,
  setKnowledgeWriter,
} from "../knowledge";

export interface KnowledgeRuntimeState {
  store: SqliteStore | null;
  embedder: EmbeddingService | null;
  indexer: KnowledgeIndexer | null;
  writer: KnowledgeWriter | null;
  retriever: Retriever | null;
  error?: string | null;
}

let knowledgeStore: SqliteStore | null = null;
let knowledgeEmbedder: EmbeddingService | null = null;
let knowledgeIndexer: KnowledgeIndexer | null = null;
let knowledgeWriter: KnowledgeWriter | null = null;
let knowledgeRetriever: Retriever | null = null;
let knowledgeRuntimeSignature: string | null = null;
let knowledgeRuntimeError: string | null = null;

/**
 * RAG 知识库运行时装配。
 *
 * 关联模块：
 * - knowledge/*: SQLite 存储、Embedding、索引与检索实现。
 * - tools/registry: 工具执行器需要 knowledgeRetriever 支持知识检索工具。
 */
export async function initializeKnowledgeRuntime(
  aiConfig: AIClientConfig,
  dataRoot?: string,
): Promise<KnowledgeRuntimeState> {
  const signature = buildKnowledgeRuntimeSignature(aiConfig, dataRoot);
  if (knowledgeStore && knowledgeRuntimeSignature !== signature) {
    resetKnowledgeRuntime();
  }

  if (!knowledgeStore) {
    try {
      const dbPath = await resolveKnowledgeDbPath(dataRoot);
      knowledgeStore = new SqliteStore(dbPath);
      await knowledgeStore.init();

      knowledgeEmbedder = new EmbeddingService({
        provider: aiConfig.provider,
        baseUrl: aiConfig.baseUrl,
        apiKey: aiConfig.apiKey,
        customHeaders: aiConfig.customHeaders,
      });

      knowledgeIndexer = new KnowledgeIndexer(knowledgeStore, knowledgeEmbedder);
      knowledgeWriter = new KnowledgeWriter(knowledgeStore, knowledgeEmbedder);
      knowledgeRetriever = new Retriever(knowledgeStore, knowledgeEmbedder);

      setKnowledgeStore(knowledgeStore);
      setKnowledgeRetriever(knowledgeRetriever);
      setKnowledgeIndexer(knowledgeIndexer);
      setKnowledgeWriter(knowledgeWriter);
      knowledgeRuntimeSignature = signature;
      knowledgeRuntimeError = null;
    } catch (e) {
      knowledgeRuntimeError = formatKnowledgeRuntimeError(e);
      clearKnowledgeRuntimeInstances();
      console.warn("RAG 知识库初始化失败（可在设置中配置后重试）:", e);
    }
  }

  return {
    store: knowledgeStore,
    embedder: knowledgeEmbedder,
    indexer: knowledgeIndexer,
    writer: knowledgeWriter,
    retriever: knowledgeRetriever,
    error: knowledgeRuntimeError,
  };
}

export function resetKnowledgeRuntime(): void {
  clearKnowledgeRuntimeInstances();
  knowledgeRuntimeError = null;
}

function clearKnowledgeRuntimeInstances(): void {
  try {
    knowledgeStore?.close();
  } catch {
    // Ignore close errors during data-path migration; the runtime will be recreated.
  }
  knowledgeStore = null;
  knowledgeEmbedder = null;
  knowledgeIndexer = null;
  knowledgeWriter = null;
  knowledgeRetriever = null;
  knowledgeRuntimeSignature = null;
  resetKnowledgeRegistry();
}

export async function reloadKnowledgeRuntime(
  aiConfig: AIClientConfig,
  dataRoot?: string,
): Promise<KnowledgeRuntimeState> {
  resetKnowledgeRuntime();
  return initializeKnowledgeRuntime(aiConfig, dataRoot);
}

async function resolveKnowledgeDbPath(dataRoot?: string): Promise<string> {
  const legacyPath = getLegacyKnowledgeDbPath();
  if (!dataRoot) return legacyPath;

  const nextPath = path.join(dataRoot, "knowledge", "knowledge.db");
  if (!fs.existsSync(nextPath) && fs.existsSync(legacyPath)) {
    await fs.promises.mkdir(path.dirname(nextPath), { recursive: true });
    await fs.promises.copyFile(legacyPath, nextPath);
  }
  return nextPath;
}

function getLegacyKnowledgeDbPath(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant", "knowledge", "knowledge.db");
}

function buildKnowledgeRuntimeSignature(aiConfig: AIClientConfig, dataRoot?: string): string {
  return JSON.stringify({
    dataRoot: dataRoot || "",
    provider: aiConfig.provider,
    baseUrl: aiConfig.baseUrl,
    apiKey: aiConfig.apiKey,
    customHeaders: aiConfig.customHeaders || {},
  });
}

function formatKnowledgeRuntimeError(error: unknown): string {
  if (error instanceof Error) return error.message;
  return String(error || "未知错误");
}
