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
    throw new Error(`${field} is not a finite number`);
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

/**
 * null = mixed/unavailable; non-null must be non-empty string (empty/whitespace rejected).
 */
export function nullableNonEmptyString(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a non-empty string or null`);
  }
  return value;
}

/**
 * null = mixed/unavailable; non-null must be finite number > 0 (fontSize, rowHeight).
 */
export function nullablePositiveFinite(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error(`${field} is not a finite number > 0 or null`);
  }
  return value;
}

/**
 * null = mixed/unavailable; non-null must be boolean.
 */
export function nullableBoolean(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a boolean or null`);
  }
  return value;
}

/** null = mixed; non-null must be exact #RRGGBB (no auto # prefix). */
export function nullableHexColor(value: unknown, field: string): string | null {
  if (value === null) return null;
  return requireHexColor(value, field);
}

/** Exact #RRGGBB only (case-insensitive compare elsewhere); no auto-prefix. */
export function requireHexColor(value: unknown, field: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a loaded color string`);
  }
  const raw = value.trim();
  if (!raw.startsWith("#") || !HEX_RE.test(raw)) {
    throw new Error(`${field} is not a #RRGGBB color`);
  }
  return raw.toUpperCase();
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
    throw new Error(`${field} is not Center`);
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

/** Strict A1 cell: letters + row >= 1 (rejects A0 / garbage). */
function parseStrictA1Cell(cell: string): { row: number; col: number } {
  const bare = cell.replace(/\$/g, "").trim();
  const match = /^([A-Za-z]+)([1-9]\d*)$/.exec(bare);
  if (!match) {
    throw new Error(`not a parseable A1 cell: ${cell}`);
  }
  const letters = match[1]!.toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  if (col < 1) throw new Error(`not a parseable A1 cell: ${cell}`);
  return { row: Number(match[2]), col };
}

/**
 * Quote-aware sheet!A1 or A1:B2 range; rejects garbage, A0, multi-area, missing endpoints.
 */
export function requireParseableA1Range(address: unknown, field: string): string {
  if (typeof address !== "string" || address.trim() === "") {
    throw new Error(`${field} is not a non-empty address string`);
  }
  const { bare } = splitSheetQualifiedAddress(address);
  if (bare.includes(",")) {
    throw new Error(`${field} multi-area address is not supported`);
  }
  const parts = bare.split(":");
  if (parts.length > 2) {
    throw new Error(`${field} is not a single cell or A1:B2 range`);
  }
  for (const part of parts) {
    if (!part || part.trim() === "") {
      throw new Error(`${field} has empty range endpoint`);
    }
    parseStrictA1Cell(part);
  }
  return bare.replace(/\$/g, "").toUpperCase();
}

/** null or non-negative integer (fit page counts). undefined is not null. */
export function nullableNonNegativeInt(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (value === undefined) {
    throw new Error(`${field} is undefined/missing`);
  }
  return requireNonNegativeInt(value, field);
}

/**
 * Orientation host token: Portrait|Landscape only (case-insensitive exact; no strip).
 * Only explicit null = unavailable; undefined fails.
 */
export function nullableOrientation(value: unknown, field: string): "portrait" | "landscape" | null {
  if (value === null) return null;
  if (value === undefined) {
    throw new Error(`${field} is undefined/missing`);
  }
  if (typeof value !== "string") {
    throw new Error(`${field} is not a string or null`);
  }
  const lower = value.toLowerCase();
  if (lower === "portrait") return "portrait";
  if (lower === "landscape") return "landscape";
  throw new Error(`${field} is not a portrait/landscape host token`);
}

/** null or non-empty paper size string (no coercion). undefined fails. */
export function nullablePaperSizeToken(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (value === undefined) {
    throw new Error(`${field} is undefined/missing`);
  }
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${field} is not a non-empty string or null`);
  }
  return value;
}

/** null or string (header/footer may be empty string). undefined fails. */
export function nullableStringAllowEmpty(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (value === undefined) {
    throw new Error(`${field} is undefined/missing`);
  }
  if (typeof value !== "string") {
    throw new Error(`${field} is not a string or null`);
  }
  return value;
}
