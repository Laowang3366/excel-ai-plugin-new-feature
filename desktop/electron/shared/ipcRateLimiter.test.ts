import { describe, expect, it } from "vitest";

import { createIpcRateLimiter } from "./ipcRateLimiter";

describe("IPC rate limiter", () => {
  it("shares a token bucket across related channels for the same sender", () => {
    const limiter = createIpcRateLimiter({
      "agent:startTurn": { bucket: "agent:turn", capacity: 2, refillPerSecond: 1 },
      "agent:enqueueTurn": { bucket: "agent:turn", capacity: 2, refillPerSecond: 1 },
    });
    const sender = { processId: 1, frameId: 2 };

    limiter.assertAllowed("agent:startTurn", sender, 1_000);
    limiter.assertAllowed("agent:enqueueTurn", sender, 1_000);
    expect(() => limiter.assertAllowed("agent:startTurn", sender, 1_000))
      .toThrow("ipc_rate_limit_exceeded:agent:startTurn");
    expect(() => limiter.assertAllowed("agent:startTurn", sender, 2_000)).not.toThrow();
  });

  it("isolates buckets by renderer sender and ignores unconfigured channels", () => {
    const limiter = createIpcRateLimiter({
      "file:writeTempFile": { bucket: "file:writeTempFile", capacity: 1, refillPerSecond: 0 },
    });

    limiter.assertAllowed("file:writeTempFile", { processId: 1, frameId: 1 }, 1_000);
    expect(() => limiter.assertAllowed(
      "file:writeTempFile",
      { processId: 2, frameId: 1 },
      1_000,
    )).not.toThrow();
    expect(() => limiter.assertAllowed(
      "settings:get",
      { processId: 1, frameId: 1 },
      1_000,
    )).not.toThrow();
  });
});
