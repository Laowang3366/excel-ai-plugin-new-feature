/** Strict scalar readback validators for workbook template apply/capture (no coercion). */

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

export function requireBoolean(value: unknown, field: string): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a loaded boolean`);
  }
  return value;
}

/** Optional style field: string or null only (mixed → null + caller limitation). */
export function optionalStringOrNull(value: unknown, field: string): string | null {
  if (value === null) return null;
  if (typeof value !== "string") {
    throw new Error(`${field} is not a string or null`);
  }
  return value;
}

export function optionalFiniteOrNull(value: unknown, field: string): number | null {
  if (value === null) return null;
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${field} is not a finite number or null`);
  }
  return value;
}

export function optionalBooleanOrNull(value: unknown, field: string): boolean | null {
  if (value === null) return null;
  if (typeof value !== "boolean") {
    throw new Error(`${field} is not a boolean or null`);
  }
  return value;
}

/** Normalize host color to #RRGGBB uppercase; reject malformed. */
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

/** Alignment: case-insensitive exact token map to Center (no punctuation strip). */
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

export function stripSheetPrefix(address: string): string {
  return address.includes("!") ? address.split("!")[1]! : address;
}
