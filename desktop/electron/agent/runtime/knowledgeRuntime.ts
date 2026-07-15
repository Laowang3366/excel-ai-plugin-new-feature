import * as path from "path";
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
import { indexBuiltinKnowledge } from "../knowledge/builtinKnowledge";
import { createLogger } from "../../shared/logger";

const knowledgeRuntimeLogger = createLogger("KnowledgeRuntime");

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
  dataRoot: string,
  isRemoteDataProcessingEnabled: () => boolean = () => false,
): Promise<KnowledgeRuntimeState> {
  const signature = buildKnowledgeRuntimeSignature(aiConfig, dataRoot);
  if (knowledgeStore && knowledgeRuntimeSignature !== signature) {
    return reloadKnowledgeRuntime(aiConfig, dataRoot, isRemoteDataProcessingEnabled);
  }

  if (!knowledgeStore) {
    try {
      const candidate = await createKnowledgeRuntime(
        aiConfig,
        dataRoot,
        isRemoteDataProcessingEnabled,
      );
      activateKnowledgeRuntime(candidate, signature);
    } catch (e) {
      knowledgeRuntimeError = formatKnowledgeRuntimeError(e);
      knowledgeRuntimeLogger.warn(
        "RAG 知识库初始化失败（可在设置中配置后重试）",
        e instanceof Error ? { message: e.message, stack: e.stack } : { error: String(e) },
      );
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
  dataRoot: string,
  isRemoteDataProcessingEnabled: () => boolean = () => false,
): Promise<KnowledgeRuntimeState> {
  const signature = buildKnowledgeRuntimeSignature(aiConfig, dataRoot);
  try {
    const candidate = await createKnowledgeRuntime(
      aiConfig,
      dataRoot,
      isRemoteDataProcessingEnabled,
    );
    activateKnowledgeRuntime(candidate, signature);
  } catch (error) {
    knowledgeRuntimeError = formatKnowledgeRuntimeError(error);
    knowledgeRuntimeLogger.warn(
      "RAG 知识库刷新失败，继续使用上一运行时",
      error instanceof Error
        ? { message: error.message, stack: error.stack }
        : { error: String(error) },
    );
  }
  return currentKnowledgeRuntimeState();
}

interface CompleteKnowledgeRuntime {
  store: SqliteStore;
  embedder: EmbeddingService;
  indexer: KnowledgeIndexer;
  writer: KnowledgeWriter;
  retriever: Retriever;
}

async function createKnowledgeRuntime(
  aiConfig: AIClientConfig,
  dataRoot: string,
  isRemoteDataProcessingEnabled: () => boolean,
): Promise<CompleteKnowledgeRuntime> {
  const dbPath = await resolveKnowledgeDbPath(dataRoot);
  const store = new SqliteStore(dbPath);
  try {
    await store.init();
    const embedder = new EmbeddingService({
      provider: aiConfig.provider,
      baseUrl: aiConfig.baseUrl,
      apiKey: aiConfig.apiKey,
      customHeaders: aiConfig.customHeaders,
      remoteDataProcessingEnabled: isRemoteDataProcessingEnabled,
    });
    const indexer = new KnowledgeIndexer(store, embedder);
    const candidate = {
      store,
      embedder,
      indexer,
      writer: new KnowledgeWriter(store, embedder),
      retriever: new Retriever(store, embedder),
    };
    await indexBuiltinKnowledge(indexer);
    return candidate;
  } catch (error) {
    try {
      store.close();
    } catch {
      /* preserve the initialization error */
    }
    throw error;
  }
}

function activateKnowledgeRuntime(candidate: CompleteKnowledgeRuntime, signature: string): void {
  const previousStore = knowledgeStore;
  knowledgeStore = candidate.store;
  knowledgeEmbedder = candidate.embedder;
  knowledgeIndexer = candidate.indexer;
  knowledgeWriter = candidate.writer;
  knowledgeRetriever = candidate.retriever;
  knowledgeRuntimeSignature = signature;
  knowledgeRuntimeError = null;

  setKnowledgeStore(candidate.store);
  setKnowledgeRetriever(candidate.retriever);
  setKnowledgeIndexer(candidate.indexer);
  setKnowledgeWriter(candidate.writer);

  if (previousStore && previousStore !== candidate.store) {
    try {
      previousStore.close();
    } catch {
      /* the replacement runtime is already active */
    }
  }
}

function currentKnowledgeRuntimeState(): KnowledgeRuntimeState {
  return {
    store: knowledgeStore,
    embedder: knowledgeEmbedder,
    indexer: knowledgeIndexer,
    writer: knowledgeWriter,
    retriever: knowledgeRetriever,
    error: knowledgeRuntimeError,
  };
}

async function resolveKnowledgeDbPath(dataRoot: string): Promise<string> {
  return path.join(dataRoot, "knowledge", "knowledge.db");
}

function buildKnowledgeRuntimeSignature(aiConfig: AIClientConfig, dataRoot: string): string {
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
