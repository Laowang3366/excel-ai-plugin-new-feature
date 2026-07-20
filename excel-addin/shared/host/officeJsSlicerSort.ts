/**
 * Slicer sortBy host mapping (case-insensitive exact structure only).
 */
import type { SlicerSortBy } from "./slicerTypes";

const SORT_TO_HOST: Record<SlicerSortBy, string> = {
  dataSourceOrder: "DataSourceOrder",
  ascending: "Ascending",
  descending: "Descending",
};

const SORT_FROM_HOST: Record<string, SlicerSortBy> = {
  datasourceorder: "dataSourceOrder",
  ascending: "ascending",
  descending: "descending",
};

export function mapSortByToHost(value: SlicerSortBy): string {
  return SORT_TO_HOST[value];
}

export function mapSortByFromHost(raw: unknown): SlicerSortBy {
  if (typeof raw !== "string") {
    throw new Error(`invalid slicer sortBy readback: ${String(raw)}`);
  }
  const mapped = SORT_FROM_HOST[raw.toLowerCase()];
  if (!mapped) {
    throw new Error(`unknown slicer sortBy host token: ${raw}`);
  }
  if (raw.toLowerCase() !== raw.toLowerCase().trim() || /\s/.test(raw)) {
    throw new Error(`unknown slicer sortBy host token: ${raw}`);
  }
  const expected = SORT_TO_HOST[mapped];
  if (raw.toLowerCase() !== expected.toLowerCase()) {
    throw new Error(`unknown slicer sortBy host token: ${raw}`);
  }
  return mapped;
}
