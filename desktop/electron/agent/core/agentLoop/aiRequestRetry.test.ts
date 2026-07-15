import { describe, expect, it, vi } from "vitest";

import {
  isRetriableAIRequestError,
  runAIRequestWithRetry,
} from "./aiRequestRetry";

describe("aiRequestRetry", () => {
  it("retries transient AI request failures before returning success", async () => {
    let attempts = 0;

    const result = await runAIRequestWithRetry({
      phase: "sampling",
      config: { maxRetries: 2, baseDelayMs: 0, maxDelayMs: 0 },
      operation: async () => {
        attempts += 1;
        if (attempts < 3) {
          const error = new Error("upstream unavailable");
          (error as any).status = 500;
          throw error;
        }
        return "ok";
      },
    });

    expect(result).toBe("ok");
    expect(attempts).toBe(3);
  });

  it("does not retry abort errors", async () => {
    let attempts = 0;
    const abortError = new Error("aborted");
    abortError.name = "AbortError";

    await expect(
      runAIRequestWithRetry({
        phase: "sampling",
        config: { maxRetries: 2, baseDelayMs: 0, maxDelayMs: 0 },
        operation: async () => {
          attempts += 1;
          throw abortError;
        },
      })
    ).rejects.toBe(abortError);
    expect(attempts).toBe(1);
  });

  it("does not retry non-transient client errors", async () => {
    let attempts = 0;
    const badRequest = new Error("bad request");
    (badRequest as any).status = 400;

    await expect(
      runAIRequestWithRetry({
        phase: "compact",
        config: { maxRetries: 2, baseDelayMs: 0, maxDelayMs: 0 },
        operation: async () => {
          attempts += 1;
          throw badRequest;
        },
      })
    ).rejects.toBe(badRequest);
    expect(attempts).toBe(1);
  });

  it("respects a caller retry guard", async () => {
    let attempts = 0;
    const upstreamError = Object.assign(new Error("upstream unavailable"), { status: 503 });

    await expect(
      runAIRequestWithRetry({
        phase: "sampling",
        config: { maxRetries: 2, baseDelayMs: 0, maxDelayMs: 0 },
        canRetry: () => false,
        operation: async () => {
          attempts += 1;
          throw upstreamError;
        },
      })
    ).rejects.toBe(upstreamError);
    expect(attempts).toBe(1);
  });

  it("classifies fetch/network failures as retriable", () => {
    expect(isRetriableAIRequestError(new TypeError("fetch failed"))).toBe(true);
    expect(isRetriableAIRequestError(new Error("read ECONNRESET"))).toBe(true);
    expect(isRetriableAIRequestError(new Error("ETIMEDOUT"))).toBe(true);
    expect(isRetriableAIRequestError(new Error("API 请求失败 (502): 模型服务网关暂时不可用"))).toBe(true);
  });

  it("waits with exponential backoff between retries", async () => {
    vi.useFakeTimers();
    try {
      const operation = vi
        .fn<() => Promise<string>>()
        .mockRejectedValueOnce(Object.assign(new Error("rate limited"), { status: 429 }))
        .mockResolvedValueOnce("ok");

      const promise = runAIRequestWithRetry({
        phase: "sampling",
        config: { maxRetries: 1, baseDelayMs: 50, maxDelayMs: 50 },
        operation,
      });

      await vi.advanceTimersByTimeAsync(49);
      expect(operation).toHaveBeenCalledTimes(1);
      await vi.advanceTimersByTimeAsync(1);
      await expect(promise).resolves.toBe("ok");
      expect(operation).toHaveBeenCalledTimes(2);
    } finally {
      vi.useRealTimers();
    }
  });
});
