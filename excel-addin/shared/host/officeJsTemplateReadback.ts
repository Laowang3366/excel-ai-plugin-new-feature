/**
 * Strict scalar/address validators for workbook template apply/capture (no coercion).
 */

const HEX_RE = /^#[0-9A-Fa-f]{6}$/;

export function requireNonEmptyString(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded non-empty string`);
  }
  return value;
}

export function requireFiniteNumber(value: unknown, field: string): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a loaded finite number`);
  }
  return value;
}

/** Strict boolean — no truthy/falsy coercion. */
export function requireBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a loaded boolean`);
  }
  return value;
}

export function requirePositiveInt(value: unknown, field: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 1
  ) {
    throw new Error(`${field} is not a positive integer`);
  }
  return value;
}

export function requireNonNegativeInt(value: unknown, field: string): number {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 0
  ) {
    throw new Error(`${field} is not a non-negative integer`);
  }
  return value;
}

/** null = mixed/unavailable; non-null must be string (bad types fail). */
export function nullableString(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (typeof value !== "string") {
    throw new Error(`${field} is not a string or null`);
  }
  return value;
}

/** null = mixed/unavailable; non-null must be finite number. */
export function nullableFiniteNumber(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a finite number or null`);
  }
  return value;
}

/** null = mixed/unavailable; non-null must be boolean. */
export function nullableBoolean(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a boolean or null`);
  }
  return value;
}

/** null = mixed; non-null must be #RRGGBB. */
export function nullableHexColor(value: unknown, field: string): string | null {
  if (value === null) return null;
  return requireHexColor(value, field);
}

export function requireHexColor(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded color string`);
  }
  const raw = value.trim();
  const withHash = raw.startsWith("#") ? raw : `#${raw}`;
  if (!HEX_RE.test(withHash)) {
    throw new Error(`${field} is not a #RRGGBB color`);
  }
  return withHash.toUpperCase();
}

export function colorsEqual(a: string, b: string): boolean {
  return a.toUpperCase() === b.toUpperCase();
}

/** Font name compare: trim + case-insensitive exact (no aliases). */
export function fontsEqual(a: string, b: string): boolean {
  return a.trim().toLowerCase() === b.trim().toLowerCase();
}

/** Alignment: case-insensitive exact token Center only. */
export function requireAlignmentCenter(value: unknown, field: string): string {
  if (typeof value !== "string") {
    throw new Error(`${field} is not a string alignment`);
  }
  if (value.toLowerCase() !== "center") {
    throw new Error(`${field} expected Center, got ${JSON.stringify(value)}`);
  }
  return "Center";
}

export function numbersClose(actual: number, expected: number, tol = 0.51): boolean {
  return Math.abs(actual - expected) <= tol;
}

/**
 * Split host address on the last `!` not inside single-quoted sheet name.
 * Handles: Sheet1!A1, 'Sheet 2'!$A$1, 'A!B'!$A$1:$C$2, ''quoted''!A1.
 */
export function splitSheetQualifiedAddress(address: string): {
  sheet: string | null;
  bare: string;
} {
  if (typeof address !== "string" || address.trim() === "") {
    throw new Error("address is not a non-empty string");
  }
  const raw = address.trim();
  let inQuotes = false;
  let lastBang = -1;
  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i]!;
    if (ch === "'") {
      if (inQuotes && raw[i + 1] === "'") {
        i += 1;
        continue;
      }
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === "!" && !inQuotes) lastBang = i;
  }
  if (lastBang < 0) {
    return { sheet: null, bare: raw.replace(/\$/g, "") };
  }
  let sheetPart = raw.slice(0, lastBang);
  const bare = raw.slice(lastBang + 1).replace(/\$/g, "");
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'") && sheetPart.length >= 2) {
    sheetPart = sheetPart.slice(1, -1).replace(/''/g, "'");
  }
  return { sheet: sheetPart, bare };
}

/** Bare A1 for compare: strip sheet prefix (quote-aware), $, uppercase. */
export function normalizeRangeAddressForCompare(address: string): string {
  const { bare } = splitSheetQualifiedAddress(address);
  if (bare.trim() === "") throw new Error("address bare range is empty");
  return bare.replace(/\$/g, "").toUpperCase();
}
