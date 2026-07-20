/**
 * Office.js slicer.filter.get / apply / clear (ExcelApi 1.10).
 * selectItems([]) = select all (official). getSelectedItems returns keys.
 * Missing stable members after 1.10 precheck → ordinary failure (not verified:false success).
 */
import {
  assertFilterSurfaceConsistent,
  parseSelectedKeys,
  parseSlicerItem,
  requireHostBoolean,
  requireHostString,
} from "./officeJsSlicerReadback";
import { queueLoadSlicer, withSlicerExcel } from "./officeJsSlicerShared";
import type { ExcelSlicer, ExcelSlicerCollection } from "./officeJsSlicerTypes";
import type {
  SlicerFilterApplyInput,
  SlicerFilterClearInput,
  SlicerFilterGetInput,
  SlicerFilterInfo,
  SlicerItemInfo,
} from "./slicerTypes";
import { SLICER_MAX_ITEMS_READBACK } from "./slicerTypes";
import type { HostResult } from "./types";

function getCollection(context: {
  workbook: { slicers?: ExcelSlicerCollection };
}): ExcelSlicerCollection {
  const coll = context.workbook.slicers;
  if (!coll || typeof coll.getItem !== "function") {
    throw new Error("Workbook.slicers.getItem is not available");
  }
  return coll;
}

async function readFilterSnapshot(
  context: { sync(): Promise<void> },
  slicer: ExcelSlicer,
  expectedKeys?: string[],
): Promise<SlicerFilterInfo> {
  queueLoadSlicer(slicer);
  if (!slicer.slicerItems || typeof slicer.slicerItems.load !== "function") {
    throw new Error("Slicer.slicerItems.load is not available");
  }
  slicer.slicerItems.load("items");
  await context.sync();

  const itemProxies = Array.isArray(slicer.slicerItems.items) ? slicer.slicerItems.items : null;
  if (!itemProxies) {
    throw new Error("Slicer.slicerItems.items must be an array after load");
  }
  for (const item of itemProxies) {
    if (typeof item.load !== "function") {
      throw new Error("SlicerItem.load is not available");
    }
    item.load("name,key,isSelected,hasData");
  }

  if (typeof slicer.getSelectedItems !== "function") {
    throw new Error("Slicer.getSelectedItems is not available");
  }
  const selectedResult = slicer.getSelectedItems();
  await context.sync();

  // ClientResult.value only valid after sync
  const selectedKeys = parseSelectedKeys(selectedResult?.value);
  const name = requireHostString("slicer.name", slicer.name);
  const isFilterCleared = requireHostBoolean("slicer.isFilterCleared", slicer.isFilterCleared);

  const allItems: SlicerItemInfo[] = itemProxies.map((item, index) => parseSlicerItem(item, index));
  const itemCount = allItems.length;
  if (itemCount > SLICER_MAX_ITEMS_READBACK) {
    throw new Error(
      `slicer has ${itemCount} items; exceeds readback cap ${SLICER_MAX_ITEMS_READBACK}`,
    );
  }

  assertFilterSurfaceConsistent(selectedKeys, allItems, isFilterCleared, expectedKeys);

  return {
    name,
    isFilterCleared,
    selectedKeys,
    items: allItems,
    itemCount,
    truncated: false,
    verified: true,
    limitations: [
      "selectItems([]) selects all items (official); empty keys is not 'select none'",
      "Stable Excel.Slicer has no sourceName/sourceType/sourceField readback",
    ],
  };
}

export async function officeJsGetSlicerFilter(
  input: SlicerFilterGetInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return withSlicerExcel("slicer.filter.get", async (context) => {
    const slicer = getCollection(context).getItem(input.name);
    return readFilterSnapshot(context, slicer);
  });
}

export async function officeJsApplySlicerFilter(
  input: SlicerFilterApplyInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return withSlicerExcel("slicer.filter.apply", async (context) => {
    const slicer = getCollection(context).getItem(input.name);
    queueLoadSlicer(slicer);
    await context.sync();
    if (typeof slicer.selectItems !== "function") {
      throw new Error("Slicer.selectItems is not available");
    }
    slicer.selectItems(input.keys);
    await context.sync();
    // keys=[] → select all: expect selectedKeys = all item keys (via surface check only)
    // non-empty → exact selectedKeys match
    const expected = input.keys.length === 0 ? undefined : input.keys;
    const snap = await readFilterSnapshot(context, slicer, expected);
    if (input.keys.length === 0 && !snap.isFilterCleared) {
      throw new Error("selectItems([]) readback: isFilterCleared must be true");
    }
    return snap;
  });
}

export async function officeJsClearSlicerFilter(
  input: SlicerFilterClearInput,
): Promise<HostResult<SlicerFilterInfo>> {
  return withSlicerExcel("slicer.filter.clear", async (context) => {
    const slicer = getCollection(context).getItem(input.name);
    queueLoadSlicer(slicer);
    await context.sync();
    if (typeof slicer.clearFilters !== "function") {
      throw new Error("Slicer.clearFilters is not available");
    }
    slicer.clearFilters();
    await context.sync();
    const snap = await readFilterSnapshot(context, slicer);
    if (snap.isFilterCleared !== true) {
      throw new Error("clearFilters readback: isFilterCleared is not true");
    }
    return snap;
  });
}
