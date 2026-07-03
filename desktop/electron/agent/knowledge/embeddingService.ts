/**
 * Embedding 服务
 *
 * 调用 AI Provider 的 /embeddings API 生成文本向量。
 * 兼容所有 OpenAI 兼容的 Embedding API（OpenAI、智谱、DeepSeek 等）。
 *
 * 设计要点：
 * - 复用与 AI Client 相同的鉴权方式（baseUrl + apiKey）
 * - 内置 LRU 文本缓存，相同文本不重复调用 API
 * - 批处理支持（同时嵌入多段文本）
 */

import { createHash } from "crypto";

// ============================================================
// 各厂商默认 embedding 模型配置
// ============================================================

interface EmbeddingModelConfig {
  model: string;
  dimensions: number;
}

export interface EmbeddingProfile {
  provider: string;
  model: string;
  dimensions: number;
}

const PROVIDER_EMBEDDING_MODELS: Record<string, EmbeddingModelConfig> = {
  openai:     { model: "text-embedding-3-small", dimensions: 1536 },
  deepseek:   { model: "deepseek-embedding",     dimensions: 1024 },
  zhipu:      { model: "embedding-2",            dimensions: 1024 },
  kimi:       { model: "text-embedding-v1",      dimensions: 1024 },
  xiaomi:     { model: "text-embedding-v1",      dimensions: 1024 },
  baidu:      { model: "embedding-v1",           dimensions: 1024 },
  aliyun:     { model: "text-embedding-v2",      dimensions: 1536 },
  volcengine: { model: "text-embedding-v1",      dimensions: 1024 },
  tencent:    { model: "text-embedding-v1",      dimensions: 1024 },
  qwen:       { model: "text-embedding-v2",      dimensions: 1536 },
  minimax:    { model: "text-embedding-v1",      dimensions: 1024 },
};

const DEFAULT_EMBEDDING_CONFIG: EmbeddingModelConfig = {
  model: "text-embedding-v1",
  dimensions: 1024,
};

// ============================================================
// LRU 缓存
// ============================================================

class LRUCache<V> {
  private cache = new Map<string, V>();
  private maxSize: number;

  constructor(maxSize = 500) {
    this.maxSize = maxSize;
  }

  get(key: string): V | undefined {
    const value = this.cache.get(key);
    if (value !== undefined) {
      // 移到最近使用
      this.cache.delete(key);
      this.cache.set(key, value);
    }
    return value;
  }

  set(key: string, value: V): void {
    if (this.cache.has(key)) {
      this.cache.delete(key);
    } else if (this.cache.size >= this.maxSize) {
      // 移除最久未使用的（Map 的第一个 key）
      const firstKey = this.cache.keys().next().value;
      if (firstKey !== undefined) this.cache.delete(firstKey);
    }
    this.cache.set(key, value);
  }

  get size(): number {
    return this.cache.size;
  }
}

// ============================================================
// EmbeddingService
// ============================================================

export interface EmbeddingServiceConfig {
  provider: string;
  baseUrl: string;
  apiKey: string;
  /** 可选覆盖默认模型名 */
  model?: string;
  /** 可选覆盖自定义请求头 */
  customHeaders?: Record<string, string>;
  /** 缓存大小（默认 500） */
  cacheSize?: number;
}

export class EmbeddingService {
  private config: EmbeddingServiceConfig;
  private modelConfig: EmbeddingModelConfig;
  private cache: LRUCache<number[]>;

  constructor(config: EmbeddingServiceConfig) {
    this.config = config;
    this.modelConfig = PROVIDER_EMBEDDING_MODELS[config.provider]
      || { model: config.model || DEFAULT_EMBEDDING_CONFIG.model, dimensions: DEFAULT_EMBEDDING_CONFIG.dimensions };
    // 如果用户显式传了 model，优先使用
    if (config.model) {
      this.modelConfig = { ...this.modelConfig, model: config.model };
    }
    this.cache = new LRUCache<number[]>(config.cacheSize || 500);
  }

  /** 获取当前使用的 embedding 模型名 */
  getModelName(): string {
    return this.modelConfig.model;
  }

  getProvider(): string {
    return this.config.provider;
  }

  getProfile(): EmbeddingProfile {
    return {
      provider: this.config.provider,
      model: this.modelConfig.model,
      dimensions: this.modelConfig.dimensions,
    };
  }

  /** 获取向量维度 */
  getDimensions(): number {
    return this.modelConfig.dimensions;
  }

  /** 获取缓存大小 */
  getCacheSize(): number {
    return this.cache.size;
  }

  /**
   * 为单段文本生成向量嵌入
   */
  async embed(text: string): Promise<number[]> {
    if (!text || text.trim().length === 0) {
      throw new Error("嵌入文本不能为空");
    }

    // 缓存命中直接返回
    const hash = this.hashText(text);
    const cached = this.cache.get(hash);
    if (cached) return cached;

    const [result] = await this.callEmbeddingAPI([text]);
    const embedding = result.embedding;

    // 写入缓存
    this.cache.set(hash, embedding);
    return embedding;
  }

  /**
   * 批量生成向量嵌入
   *
   * 先检查缓存，只对未缓存文本调用 API。
   */
  async embedBatch(texts: string[]): Promise<number[][]> {
    if (texts.length === 0) return [];

    // 检查缓存
    const results: (number[] | null)[] = new Array(texts.length).fill(null);
    const uncachedIndices: number[] = [];
    const uncachedTexts: string[] = [];

    for (let i = 0; i < texts.length; i++) {
      const text = texts[i];
      if (!text || text.trim().length === 0) {
        throw new Error(`嵌入文本不能为空（索引 ${i}）`);
      }
      const hash = this.hashText(text);
      const cached = this.cache.get(hash);
      if (cached) {
        results[i] = cached;
      } else {
        uncachedIndices.push(i);
        uncachedTexts.push(text);
      }
    }

    // 批量调用 API
    if (uncachedTexts.length > 0) {
      const apiResults = await this.callEmbeddingAPI(uncachedTexts);
      for (let j = 0; j < uncachedTexts.length; j++) {
        const { embedding } = apiResults[j];
        const text = uncachedTexts[j];
        const hash = this.hashText(text);
        this.cache.set(hash, embedding);
        results[uncachedIndices[j]] = embedding;
      }
    }

    return results as number[][];
  }

  /** 清空缓存 */
  clearCache(): void {
    // 通过重新创建实现
    this.cache = new LRUCache<number[]>(this.config.cacheSize || 500);
  }

  // ============================================================
  // 内部实现
  // ============================================================

  /**
   * 调用 Embedding API
   * 兼容 OpenAI 格式：POST /v1/embeddings
   * 请求：{ input: string | string[], model: string }
   * 响应：{ data: [{ embedding: number[] }] }
   */
  private async callEmbeddingAPI(
    texts: string[]
  ): Promise<Array<{ embedding: number[] }>> {
    const url = buildEmbeddingUrl(this.config.baseUrl);

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${this.config.apiKey}`,
      ...this.config.customHeaders,
    };

    const body = JSON.stringify({
      input: texts,
      model: this.modelConfig.model,
    });

    const response = await fetch(url, {
      method: "POST",
      headers,
      body,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Embedding API 请求失败 (${response.status}): ${formatApiError(errorText)}`
      );
    }

    const json: any = await response.json();

    // 解析响应，兼容不同格式
    if (json.data && Array.isArray(json.data)) {
      // OpenAI 标准格式: { data: [{ embedding: [...] }] }
      return json.data.map((item: any) => ({
        embedding: item.embedding as number[],
      }));
    }

    throw new Error(`Embedding API 返回格式异常: ${JSON.stringify(json).slice(0, 200)}`);
  }

  /** 计算文本的 SHA256 哈希（前 16 字符作为缓存 key） */
  private hashText(text: string): string {
    return createHash("sha256").update(text, "utf-8").digest("hex").slice(0, 16);
  }
}

function buildEmbeddingUrl(baseUrl: string): string {
  const trimmed = baseUrl.replace(/\/+$/, "");
  if (/\/embeddings$/i.test(trimmed)) return trimmed;
  if (/\/chat\/completions$/i.test(trimmed)) {
    return trimmed.replace(/\/chat\/completions$/i, "/embeddings");
  }
  if (/\/v\d+(?:\/[^/]*)?$/i.test(trimmed)) return `${trimmed}/embeddings`;
  return `${trimmed}/v1/embeddings`;
}

function formatApiError(errorText: string): string {
  const compact = errorText.replace(/\s+/g, " ").trim();
  if (!compact) return "空响应";
  const withoutHtml = compact.replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim();
  return (withoutHtml || compact).slice(0, 300);
}
