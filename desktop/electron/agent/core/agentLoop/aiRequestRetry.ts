/**
 * AI 请求重试策略。
 *
 * 关联模块：
 * - agentLoop.ts: 对 sampling 与 compact 两类模型请求应用独立策略。
 * - streamCollector.ts: 将可重试的流内 API 错误抛出，交给本模块统一处理。
 */

export type AIRequestPhase = "sampling" | "compact";

export interface AIRequestRetryConfig {
  /** 失败后的重试次数，不包含首次请求。 */
  maxRetries?: number;
  /** 首次重试等待时间。 */
  baseDelayMs?: number;
  /** 单次等待时间上限。 */
  maxDelayMs?: number;
  /** 指数退避倍率。 */
  backoffFactor?: number;
}

export const DEFAULT_SAMPLING_RETRY_CONFIG: Required<AIRequestRetryConfig> = {
  maxRetries: 2,
  baseDelayMs: 500,
  maxDelayMs: 4_000,
  backoffFactor: 2,
};

export const DEFAULT_COMPACT_RETRY_CONFIG: Required<AIRequestRetryConfig> = {
  maxRetries: 2,
  baseDelayMs: 800,
  maxDelayMs: 6_000,
  backoffFactor: 2,
};

export async function runAIRequestWithRetry<T>(params: {
  phase: AIRequestPhase;
  config?: AIRequestRetryConfig;
  signal?: AbortSignal;
  canRetry?: (error: unknown) => boolean;
  operation: () => Promise<T>;
}): Promise<T> {
  const retryConfig = normalizeRetryConfig(params.phase, params.config);
  let lastError: unknown;

  for (let attempt = 0; attempt <= retryConfig.maxRetries; attempt++) {
    if (params.signal?.aborted) throw createAbortError();
    try {
      return await params.operation();
    } catch (error) {
      lastError = error;
      if (
        attempt >= retryConfig.maxRetries ||
        params.signal?.aborted ||
        !isRetriableAIRequestError(error) ||
        params.canRetry?.(error) === false
      ) {
        throw error;
      }
      await sleep(computeDelayMs(attempt, retryConfig), params.signal);
    }
  }

  throw lastError instanceof Error ? lastError : new Error(String(lastError));
}

export function isRetriableAIRequestError(error: unknown): boolean {
  if (isAbortError(error)) return false;

  const status = getHttpStatus(error);
  if (status !== undefined) {
    return status === 408 || status === 409 || status === 425 || status === 429 || status >= 500;
  }

  const message = getErrorMessage(error).toLowerCase();
  return [
    "fetch failed",
    "networkerror",
    "econnreset",
    "econnrefused",
    "etimedout",
    "socket hang up",
    "timeout",
  ].some((needle) => message.includes(needle));
}

function normalizeRetryConfig(
  phase: AIRequestPhase,
  config?: AIRequestRetryConfig,
): Required<AIRequestRetryConfig> {
  const defaults =
    phase === "compact" ? DEFAULT_COMPACT_RETRY_CONFIG : DEFAULT_SAMPLING_RETRY_CONFIG;
  const maxRetries = Math.max(0, Math.floor(config?.maxRetries ?? defaults.maxRetries));
  const baseDelayMs = Math.max(0, Math.floor(config?.baseDelayMs ?? defaults.baseDelayMs));
  const maxDelayMs = Math.max(baseDelayMs, Math.floor(config?.maxDelayMs ?? defaults.maxDelayMs));
  const backoffFactor = Math.max(1, config?.backoffFactor ?? defaults.backoffFactor);
  return { maxRetries, baseDelayMs, maxDelayMs, backoffFactor };
}

function computeDelayMs(attempt: number, config: Required<AIRequestRetryConfig>): number {
  const delay = config.baseDelayMs * Math.pow(config.backoffFactor, attempt);
  return Math.min(config.maxDelayMs, Math.floor(delay));
}

function sleep(ms: number, signal?: AbortSignal): Promise<void> {
  if (ms <= 0) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(resolve, ms);
    signal?.addEventListener(
      "abort",
      () => {
        clearTimeout(timer);
        reject(createAbortError());
      },
      { once: true },
    );
  });
}

function getHttpStatus(error: unknown): number | undefined {
  const anyError = error as any;
  const status = anyError?.status ?? anyError?.statusCode ?? anyError?.response?.status;
  if (typeof status === "number") return status;

  const match = getErrorMessage(error).match(/\((\d{3})\)|\bstatus\s*(\d{3})\b/i);
  const fromMessage = match?.[1] ?? match?.[2];
  return fromMessage ? Number(fromMessage) : undefined;
}

function getErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error ?? "");
}

function isAbortError(error: unknown): boolean {
  return (error as { name?: string } | null)?.name === "AbortError";
}

function createAbortError(): Error {
  const error = new Error("aborted");
  error.name = "AbortError";
  return error;
}
