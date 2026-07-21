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

/**
 * Normalize host A1 for selection surfaces shared with Office.js-style contracts:
 * strip absolute `$` only from A1 reference parts; keep sheet qualifiers (incl. `$` /
 * `!` / `,` inside single quotes and doubled apostrophes). Multi-area commas outside
 * quotes are preserved as area separators.
 * Examples: `$G$17` → `G17`, `'Budget$2026'!$A$1` → `'Budget$2026'!A1`.
 */
export function normalizeWpsA1Address(address: string): string {
  if (typeof address !== "string") return address;
  return splitOutsideSingleQuotes(address, ",")
    .map(normalizeWpsA1Area)
    .join(",");
}

function normalizeWpsA1Area(area: string): string {
  const bang = lastIndexOutsideSingleQuotes(area, "!");
  if (bang < 0) {
    return area.replace(/\$/g, "");
  }
  return area.slice(0, bang + 1) + area.slice(bang + 1).replace(/\$/g, "");
}

/** Last index of `target` not inside a single-quoted sheet name (`''` = escaped). */
function lastIndexOutsideSingleQuotes(s: string, target: string): number {
  let inQuotes = false;
  let last = -1;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i]!;
    if (ch === "'") {
      if (inQuotes && s[i + 1] === "'") {
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === target && !inQuotes) last = i;
  }
  return last;
}

/** Split on `sep` only outside single-quoted sheet names. */
function splitOutsideSingleQuotes(s: string, sep: string): string[] {
  const parts: string[] = [];
  let start = 0;
  let inQuotes = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i]!;
    if (ch === "'") {
      if (inQuotes && s[i + 1] === "'") {
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === sep && !inQuotes) {
      parts.push(s.slice(start, i));
      start = i + 1;
    }
  }
  parts.push(s.slice(start));
  return parts;
}
