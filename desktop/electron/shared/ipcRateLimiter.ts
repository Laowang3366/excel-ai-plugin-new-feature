export interface IpcRateLimitPolicy {
  bucket: string;
  capacity: number;
  refillPerSecond: number;
}

export interface IpcRateLimitSender {
  processId: number;
  frameId: number;
}

interface IpcRateLimitBucket {
  tokens: number;
  updatedAt: number;
}

export const IPC_RATE_LIMIT_POLICIES: Readonly<Record<string, IpcRateLimitPolicy>> = {
  "agent:startTurn": { bucket: "agent:turn", capacity: 8, refillPerSecond: 0.5 },
  "agent:enqueueTurn": { bucket: "agent:turn", capacity: 8, refillPerSecond: 0.5 },
  "app:log": { bucket: "app:log", capacity: 200, refillPerSecond: 50 },
  "excel:writeRange": { bucket: "excel:writeRange", capacity: 30, refillPerSecond: 2 },
  "file:readAsBase64": { bucket: "file:readAsBase64", capacity: 20, refillPerSecond: 1 },
  "file:writeTempFile": { bucket: "file:writeTempFile", capacity: 10, refillPerSecond: 0.5 },
  "ocr:recognize": { bucket: "ocr:recognize", capacity: 4, refillPerSecond: 0.1 },
  "office:automation:templates:run": {
    bucket: "office:automation:templates:run",
    capacity: 10,
    refillPerSecond: 0.2,
  },
  "settings:set": { bucket: "settings:set", capacity: 100, refillPerSecond: 20 },
};

export function createIpcRateLimiter(
  policies: Readonly<Record<string, IpcRateLimitPolicy>> = IPC_RATE_LIMIT_POLICIES,
) {
  const buckets = new Map<string, IpcRateLimitBucket>();

  return {
    assertAllowed(channel: string, sender: IpcRateLimitSender, now = Date.now()): void {
      const policy = policies[channel];
      if (!policy) return;

      const key = `${sender.processId}:${sender.frameId}:${policy.bucket}`;
      const previous = buckets.get(key) ?? {
        tokens: policy.capacity,
        updatedAt: now,
      };
      const elapsedSeconds = Math.max(0, now - previous.updatedAt) / 1_000;
      const tokens = Math.min(
        policy.capacity,
        previous.tokens + elapsedSeconds * policy.refillPerSecond,
      );

      if (tokens < 1) {
        buckets.set(key, { tokens, updatedAt: now });
        throw new Error(`ipc_rate_limit_exceeded:${channel}`);
      }

      buckets.set(key, { tokens: tokens - 1, updatedAt: now });
      if (buckets.size > 256) {
        const staleBefore = now - 10 * 60 * 1_000;
        for (const [bucketKey, bucket] of buckets) {
          if (bucket.updatedAt < staleBefore) buckets.delete(bucketKey);
        }
      }
    },
    clear(): void {
      buckets.clear();
    },
  };
}
