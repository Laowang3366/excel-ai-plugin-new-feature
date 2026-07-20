/**
 * Fail-closed concurrency and token-bucket style rate limiting.
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

export class RateLimiter {
  /**
   * @param {number} max
   * @param {number} windowMs
   */
  constructor(max, windowMs) {
    if (!Number.isInteger(max) || max <= 0) {
      throw new Error("rate limit max must be a positive integer");
    }
    if (!Number.isInteger(windowMs) || windowMs <= 0) {
      throw new Error("rate limit window must be a positive integer");
    }
    this.max = max;
    this.windowMs = windowMs;
    /** @type {Map<string, number[]>} */
    this.hits = new Map();
  }

  /**
   * @param {string} key
   * @param {number} [now]
   */
  allow(key, now = Date.now()) {
    const cutoff = now - this.windowMs;
    const list = this.hits.get(key) || [];
    const kept = list.filter((t) => t > cutoff);
    if (kept.length >= this.max) {
      this.hits.set(key, kept);
      return false;
    }
    kept.push(now);
    this.hits.set(key, kept);
    return true;
  }
}
