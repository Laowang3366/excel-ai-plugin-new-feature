/**
 * Strict pure Slicer/SlicerItem readback validators (no String/Number/Boolean coercion).
 * Bad host surface → throw ordinary failure (never limitation+success).
 */
import type { ExcelSlicer, ExcelSlicerItem } from "./officeJsSlicerTypes";
import type { SlicerInfo, SlicerItemInfo, SlicerSortBy } from "./slicerTypes";
import { mapSortByFromHost } from "./officeJsSlicerSort";

export function requireHostString(label: string, value: unknown, opts?: { allowEmpty?: boolean }): string {
  if (typeof value !== "string") {
    throw new Error(`${label} must be a string (got ${typeof value})`);
  }
  if (!opts?.allowEmpty && value === "") {
    throw new Error(`${label} must be a non-empty string`);
  }
  return value;
}

export function requireHostFiniteNumber(label: string, value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${label} must be a finite number (got ${typeof value})`);
  }
  return value;
}

export function requireHostBoolean(label: string, value: unknown): boolean {
  if (typeof value !== "boolean") {
    throw new Error(`${label} must be a boolean (got ${typeof value})`);
  }
  return value;
}

export function parseSelectedKeys(raw: unknown): string[] {
  if (!Array.isArray(raw)) {
    throw new Error(`selectedKeys must be a string array (got ${typeof raw})`);
  }
  const keys: string[] = [];
  const seen = new Set<string>();
  for (let i = 0; i < raw.length; i += 1) {
    const item = raw[i];
    if (typeof item !== "string" || item === "") {
      throw new Error(`selectedKeys[${i}] must be a non-empty string`);
    }
    if (seen.has(item)) {
      throw new Error(`selectedKeys contains duplicate: ${item}`);
    }
    seen.add(item);
    keys.push(item);
  }
  return keys;
}

export function parseSlicerItem(item: ExcelSlicerItem, index: number): SlicerItemInfo {
  const prefix = `slicerItems[${index}]`;
  return {
    key: requireHostString(`${prefix}.key`, item.key),
    name: requireHostString(`${prefix}.name`, item.name, { allowEmpty: true }),
    isSelected: requireHostBoolean(`${prefix}.isSelected`, item.isSelected),
    hasData: requireHostBoolean(`${prefix}.hasData`, item.hasData),
  };
}

/**
 * Strict scalar snapshot from a loaded Slicer ClientObject.
 * limitations only for non-surface policy notes (e.g. no source readback), not bad types.
 */
export function readSlicerSnapshotStrict(
  slicer: ExcelSlicer,
  limitations: string[] = [],
): SlicerInfo {
  const name = requireHostString("slicer.name", slicer.name);
  const id = requireHostString("slicer.id", slicer.id);
  const caption = requireHostString("slicer.caption", slicer.caption, { allowEmpty: true });
  const style = requireHostString("slicer.style", slicer.style);
  const sheetName = requireHostString("slicer.worksheet.name", slicer.worksheet?.name);
  const top = requireHostFiniteNumber("slicer.top", slicer.top);
  const left = requireHostFiniteNumber("slicer.left", slicer.left);
  const width = requireHostFiniteNumber("slicer.width", slicer.width);
  const height = requireHostFiniteNumber("slicer.height", slicer.height);
  if (top < 0) throw new Error("slicer.top must be >= 0");
  if (left < 0) throw new Error("slicer.left must be >= 0");
  if (width <= 0) throw new Error("slicer.width must be > 0");
  if (height <= 0) throw new Error("slicer.height must be > 0");
  const isFilterCleared = requireHostBoolean("slicer.isFilterCleared", slicer.isFilterCleared);
  const sortBy: SlicerSortBy = mapSortByFromHost(slicer.sortBy);

  const info: SlicerInfo = {
    name,
    id,
    caption,
    sheetName,
    top,
    left,
    width,
    height,
    sortBy,
    style,
    isFilterCleared,
  };
  if (limitations.length > 0) info.limitations = [...limitations];
  return info;
}

export function setsEqual(a: readonly string[], b: readonly string[]): boolean {
  if (a.length !== b.length) return false;
  const set = new Set(a);
  return b.every((k) => set.has(k));
}

/** Reject duplicate SlicerItem.key values. */
export function assertUniqueItemKeys(items: SlicerItemInfo[]): void {
  const seen = new Set<string>();
  for (let i = 0; i < items.length; i += 1) {
    const key = items[i]!.key;
    if (seen.has(key)) {
      throw new Error(`slicerItems has duplicate key: ${key}`);
    }
    seen.add(key);
  }
}

/**
 * Full filter surface consistency:
 * - unique item keys
 * - selectedKeys ⊆ item keys, no unknown keys
 * - each item.isSelected matches membership in selectedKeys
 * - isFilterCleared === allSelected (all items selected and selectedKeys covers all keys)
 */
export function assertFilterSurfaceConsistent(
  selectedKeys: string[],
  items: SlicerItemInfo[],
  isFilterCleared: boolean,
  expectedKeys?: string[],
): { allSelected: boolean } {
  assertUniqueItemKeys(items);
  const itemKeys = items.map((i) => i.key);
  const itemKeySet = new Set(itemKeys);
  for (const key of selectedKeys) {
    if (!itemKeySet.has(key)) {
      throw new Error(`selectedKeys contains unknown item key: ${key}`);
    }
  }
  for (const item of items) {
    const should = selectedKeys.includes(item.key);
    if (item.isSelected !== should) {
      throw new Error(
        `item isSelected mismatch for key=${item.key}: isSelected=${item.isSelected} selectedKeysHas=${should}`,
      );
    }
  }
  const allSelected =
    items.length > 0 &&
    items.every((i) => i.isSelected) &&
    setsEqual(selectedKeys, itemKeys);
  // Empty slicer: treat allSelected as isFilterCleared must still match (both true when no items)
  const allSelectedOrEmpty = items.length === 0 ? true : allSelected;
  if (isFilterCleared !== allSelectedOrEmpty) {
    throw new Error(
      `isFilterCleared=${isFilterCleared} inconsistent with selection (allSelected=${allSelectedOrEmpty})`,
    );
  }
  if (expectedKeys !== undefined) {
    if (!setsEqual(selectedKeys, expectedKeys)) {
      throw new Error(
        `selectedKeys mismatch: requested=[${expectedKeys.join(",")}] got=[${selectedKeys.join(",")}]`,
      );
    }
  }
  return { allSelected: allSelectedOrEmpty };
}

/** @deprecated use assertFilterSurfaceConsistent */
export function assertSelectionConsistent(
  selectedKeys: string[],
  items: SlicerItemInfo[],
  mode: "exact-keys" | "all-selected",
  expectedKeys?: string[],
): void {
  const isFilterCleared = mode === "all-selected";
  assertFilterSurfaceConsistent(
    selectedKeys,
    items,
    isFilterCleared,
    mode === "all-selected" ? undefined : expectedKeys,
  );
  if (mode === "all-selected") {
    const itemKeys = items.map((i) => i.key);
    if (!items.every((i) => i.isSelected) || !setsEqual(selectedKeys, itemKeys)) {
      throw new Error("expected all slicer items selected");
    }
  }
}
