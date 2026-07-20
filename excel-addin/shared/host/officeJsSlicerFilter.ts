/**
 * Office.js slicer.filter.get / apply / clear (ExcelApi 1.10).
 * selectItems([]) = select all (official). getSelectedItems returns keys.
 */
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
): Promise<SlicerFilterInfo> {
  queueLoadSlicer(slicer);
  if (!slicer.slicerItems || typeof slicer.slicerItems.load !== "function") {
    throw new Error("Slicer.slicerItems is not available");
  }
  slicer.slicerItems.load("items");
  await context.sync();

  const itemProxies = Array.isArray(slicer.slicerItems.items) ? slicer.slicerItems.items : [];
  for (const item of itemProxies) {
    if (typeof item.load === "function") item.load("name,key,isSelected,hasData");
  }

  let selectedKeys: string[] = [];
  let selectedVerified = true;
  if (typeof slicer.getSelectedItems === "function") {
    const result = slicer.getSelectedItems();
    await context.sync();
    const value = result?.value;
    if (!Array.isArray(value)) {
      selectedVerified = false;
      selectedKeys = [];
    } else {
      selectedKeys = value.map((k) => String(k));
    }
  } else {
    selectedVerified = false;
  }
  await context.sync();

  const allItems: SlicerItemInfo[] = itemProxies.map((item) => ({
    key: String(item.key ?? ""),
    name: String(item.name ?? ""),
    isSelected: Boolean(item.isSelected),
    hasData: Boolean(item.hasData),
  }));
  const itemCount = allItems.length;
  const truncated = itemCount > SLICER_MAX_ITEMS_READBACK;
  const items = truncated ? allItems.slice(0, SLICER_MAX_ITEMS_READBACK) : allItems;

  const limitations: string[] = [];
  if (!selectedVerified) {
    limitations.push("getSelectedItems unavailable or non-array; selectedKeys not host-verified");
  }
  if (truncated) {
    limitations.push(`items truncated to ${SLICER_MAX_ITEMS_READBACK} of ${itemCount}`);
  }
  limitations.push(
    "selectItems([]) selects all items (official); empty keys is not 'select none'",
  );

  return {
    name: String(slicer.name),
    isFilterCleared: Boolean(slicer.isFilterCleared),
    selectedKeys,
    items,
    itemCount,
    truncated,
    verified: selectedVerified,
    limitations,
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
    // Official: empty array selects all.
    slicer.selectItems(input.keys);
    await context.sync();
    const snap = await readFilterSnapshot(context, slicer);
    if (input.keys.length === 0) {
      // Select-all: prefer isFilterCleared true when host reports it.
      if (snap.isFilterCleared !== true && snap.verified) {
        // Not all hosts flip isFilterCleared on select-all; keep verified on keys.
        snap.limitations = [
          ...(snap.limitations ?? []),
          "select-all applied; isFilterCleared may remain false depending on host",
        ];
      }
      return snap;
    }
    if (snap.verified) {
      const want = new Set(input.keys);
      const got = new Set(snap.selectedKeys);
      const match =
        want.size === got.size && [...want].every((k) => got.has(k));
      if (!match) {
        throw new Error(
          `selectItems readback mismatch: requested=${input.keys.join(",")} got=${snap.selectedKeys.join(",")}`,
        );
      }
    } else {
      snap.verified = false;
      snap.limitations = [
        ...(snap.limitations ?? []),
        "could not verify selected keys after selectItems",
      ];
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
