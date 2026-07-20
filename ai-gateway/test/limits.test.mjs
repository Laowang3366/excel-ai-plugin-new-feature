import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { RateLimiter } from "../src/limits.mjs";

describe("rate limiter bounds", () => {
  it("fail-closed when key table is saturated", () => {
    const rl = new RateLimiter(100, 60_000, { maxKeys: 2 });
    assert.equal(rl.allow("a"), true);
    assert.equal(rl.allow("b"), true);
    assert.equal(rl.allow("c"), false);
  });

  it("prunes expired keys", () => {
    const rl = new RateLimiter(2, 50, { maxKeys: 10 });
    assert.equal(rl.allow("a", 1000), true);
    assert.equal(rl.allow("b", 1000), true);
    // after window, keys pruned and new key accepted
    assert.equal(rl.allow("c", 1060), true);
  });
});
