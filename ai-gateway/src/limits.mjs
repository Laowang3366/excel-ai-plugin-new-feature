/**
 * Fail-closed concurrency and rate limiting with bounded key storage.
 */

export class ConcurrencyLimiter {
  /**
   * @param {number} max
   */
  constructor(max) {
    if (!Number.isInteger(max) || max <= 0) {
      throw new Error("max concurrent must be a positive integer");
    }
    this.max = max;
    this.active = 0;
  }

  tryAcquire() {
    if (this.active >= this.max) return false;
    this.active += 1;
    return true;
  }

  release() {
    if (this.active > 0) this.active -= 1;
  }
}

const DEFAULT_MAX_KEYS = 10_000;

export class RateLimiter {
  /**
   * @param {number} max
   * @param {number} windowMs
   * @param {{ maxKeys?: number }} [opts]
   */
  constructor(max, windowMs, opts = {}) {
    if (!Number.isInteger(max) || max <= 0) {
      throw new Error("rate limit max must be a positive integer");
    }
    if (!Number.isInteger(windowMs) || windowMs <= 0) {
      throw new Error("rate limit window must be a positive integer");
    }
    this.max = max;
    this.windowMs = windowMs;
    this.maxKeys = opts.maxKeys ?? DEFAULT_MAX_KEYS;
    /** @type {Map<string, number[]>} */
    this.hits = new Map();
  }

  /**
   * @param {string} key
   * @param {number} [now]
   */
  allow(key, now = Date.now()) {
    const safeKey = typeof key === "string" && key.length <= 128 ? key : "invalid";
    this.prune(now);
    const cutoff = now - this.windowMs;
    const list = this.hits.get(safeKey) || [];
    const kept = list.filter((t) => t > cutoff);
    if (kept.length >= this.max) {
      this.hits.set(safeKey, kept);
      return false;
    }
    if (!this.hits.has(safeKey) && this.hits.size >= this.maxKeys) {
      // Fail-closed when key table is saturated (e.g. random XFF flood).
      return false;
    }
    kept.push(now);
    this.hits.set(safeKey, kept);
    return true;
  }

  /**
   * @param {number} now
   */
  prune(now) {
    const cutoff = now - this.windowMs;
    for (const [key, list] of this.hits) {
      const kept = list.filter((t) => t > cutoff);
      if (kept.length === 0) this.hits.delete(key);
      else this.hits.set(key, kept);
    }
  }
}
