/**
 * Pure matrix/format normalization helpers for Office.js host adapters.
 */
import type { CellValue } from "./types";

export function normalizeMatrix(values: unknown): CellValue[][] {
  if (!Array.isArray(values)) return [];
  return values.map((row) => {
    if (!Array.isArray(row)) return [row as CellValue];
    return row.map((cell) => (cell === undefined ? null : (cell as CellValue)));
  });
}

export function normalizeFormulas(values: unknown): string[][] {
  if (!Array.isArray(values)) return [];
  return values.map((row) => {
    if (!Array.isArray(row)) return [String(row ?? "")];
    return row.map((cell) => String(cell ?? ""));
  });
}

export function firstNumberFormat(value: unknown): string | null {
  if (typeof value === "string") return value;
  if (!Array.isArray(value) || value.length === 0) return null;
  const first = value[0];
  if (Array.isArray(first)) return first[0] != null ? String(first[0]) : null;
  return first != null ? String(first) : null;
}

