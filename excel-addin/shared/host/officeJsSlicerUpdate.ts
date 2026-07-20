/**
 * Office.js slicer.update — stable writable fields only (ExcelApi 1.10).
 */
import {
  applyWritableSlicerFields,
  queueLoadSlicer,
  readSlicerSnapshot,
  withSlicerExcel,
} from "./officeJsSlicerShared";
import type { ExcelSlicerCollection } from "./officeJsSlicerTypes";
import type { SlicerInfo, SlicerUpdateInput } from "./slicerTypes";
import type { HostResult } from "./types";
import { fail } from "./types";

function hasUpdateFields(input: SlicerUpdateInput): boolean {
  return (
    input.newName !== undefined ||
    input.caption !== undefined ||
    input.top !== undefined ||
    input.left !== undefined ||
    input.width !== undefined ||
    input.height !== undefined ||
    input.style !== undefined ||
    input.sortBy !== undefined
  );
}

export async function officeJsUpdateSlicer(
  input: SlicerUpdateInput,
): Promise<HostResult<SlicerInfo>> {
  if (!hasUpdateFields(input)) {
    return fail("slicer.update", "office-js", "empty update: provide at least one writable field");
  }

  return withSlicerExcel("slicer.update", async (context) => {
    const coll = context.workbook.slicers as ExcelSlicerCollection;
    if (!coll || typeof coll.getItem !== "function") {
      throw new Error("Workbook.slicers.getItem is not available");
    }
    const slicer = coll.getItem(input.name);
    queueLoadSlicer(slicer);
    await context.sync();

    applyWritableSlicerFields(slicer, {
      name: input.newName,
      caption: input.caption,
      top: input.top,
      left: input.left,
      width: input.width,
      height: input.height,
      style: input.style,
      sortBy: input.sortBy,
    });
    await context.sync();
    queueLoadSlicer(slicer);
    await context.sync();

    const snapshot = readSlicerSnapshot(slicer, [
      "Stable Excel.Slicer has no source readback; update only verifies writable scalars",
    ]);
    if (input.newName !== undefined && snapshot.name !== input.newName) {
      throw new Error(`name readback mismatch: ${snapshot.name}`);
    }
    if (input.caption !== undefined && snapshot.caption !== input.caption) {
      throw new Error(`caption readback mismatch: ${snapshot.caption}`);
    }
    if (input.top !== undefined && snapshot.top !== input.top) {
      throw new Error(`top readback mismatch: ${snapshot.top}`);
    }
    if (input.left !== undefined && snapshot.left !== input.left) {
      throw new Error(`left readback mismatch: ${snapshot.left}`);
    }
    if (input.width !== undefined && snapshot.width !== input.width) {
      throw new Error(`width readback mismatch: ${snapshot.width}`);
    }
    if (input.height !== undefined && snapshot.height !== input.height) {
      throw new Error(`height readback mismatch: ${snapshot.height}`);
    }
    if (input.style !== undefined && snapshot.style !== input.style) {
      throw new Error(`style readback mismatch: ${snapshot.style}`);
    }
    if (input.sortBy !== undefined && snapshot.sortBy !== input.sortBy) {
      throw new Error(`sortBy readback mismatch: ${snapshot.sortBy}`);
    }
    return snapshot;
  });
}
