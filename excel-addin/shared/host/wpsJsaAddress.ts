/**
 * WPS JSA Range/Selection Address reader.
 *
 * Real WPS (e.g. 12.1.0.26885) exposes Address as a zero-arg method on some
 * Range/Selection objects, not only a string property. Coercing with String()
 * yields "function Address..." and must never be treated as a cell address.
 */

export type WpsAddressOwner = {
  Address?: unknown;
} | null | undefined;

function normalizeFallback(fallback?: string | null): string | undefined {
  if (fallback == null) return undefined;
  if (typeof fallback !== "string") return undefined;
  const trimmed = fallback.trim();
  return trimmed === "" ? undefined : trimmed;
}

/**
 * Read a usable A1 address from a WPS range-like owner.
 * - string property (non-empty after trim) → returned as-is (trimmed)
 * - zero-arg function → called with owner as `this`
 * - never uses Function.prototype.toString / String(fn)
 * - on missing member, throw, non-string, empty → fallback if provided, else undefined
 */
export function readWpsAddress(
  owner: WpsAddressOwner,
  fallback?: string | null,
): string | undefined {
  const fb = normalizeFallback(fallback);
  if (owner == null || (typeof owner !== "object" && typeof owner !== "function")) {
    return fb;
  }

  let raw: unknown;
  try {
    raw = (owner as { Address?: unknown }).Address;
  } catch {
    return fb;
  }

  if (typeof raw === "function") {
    try {
      raw = (raw as (this: unknown, ...args: unknown[]) => unknown).call(owner);
    } catch {
      return fb;
    }
  }

  // Reject residual callables / objects — never String(fn).
  if (typeof raw === "function" || (raw != null && typeof raw === "object")) {
    return fb;
  }
  if (typeof raw !== "string") {
    return fb;
  }
  const trimmed = raw.trim();
  if (trimmed === "") {
    return fb;
  }
  // Guard accidental function source leakage (String(fn) / method shorthand).
  if (
    /^function\b/i.test(trimmed) ||
    /^Address\s*\(/i.test(trimmed) ||
    trimmed.includes("\n") && /\breturn\b/.test(trimmed) && /\{/.test(trimmed)
  ) {
    return fb;
  }
  return trimmed;
}

/** True when owner exposes Address as a non-empty string or as a function method. */
export function hasWpsAddressSurface(owner: WpsAddressOwner): boolean {
  if (owner == null || (typeof owner !== "object" && typeof owner !== "function")) {
    return false;
  }
  try {
    const raw = (owner as { Address?: unknown }).Address;
    if (typeof raw === "function") return true;
    if (typeof raw === "string") return raw.trim() !== "";
    return false;
  } catch {
    return false;
  }
}
