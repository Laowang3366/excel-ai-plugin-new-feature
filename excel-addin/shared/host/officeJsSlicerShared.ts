/**
 * Shared Slicer runner and writable field apply (ExcelApi 1.10).
 */
import { requireExcelApi110ForSlicer } from "./officeJsSlicerRequirements";
import type { ExcelRequestContextWithSlicer, ExcelSlicer } from "./officeJsSlicerTypes";
import type { SlicerSortBy } from "./slicerTypes";
import { mapSortByToHost } from "./officeJsSlicerSort";
import { getExcelRun } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

export const SLICER_SCALAR_PROPS =
  "name,id,caption,height,width,top,left,sortBy,style,isFilterCleared";

export { mapSortByFromHost, mapSortByToHost } from "./officeJsSlicerSort";
export { readSlicerSnapshotStrict as readSlicerSnapshot } from "./officeJsSlicerReadback";

export async function withSlicerExcel<T>(
  capability: string,
  fn: (context: ExcelRequestContextWithSlicer) => Promise<T>,
): Promise<HostResult<T>> {
  const gate = requireExcelApi110ForSlicer(capability);
  if (gate) return gate as HostResult<T>;
  const run = getExcelRun();
  if (!run) {
    return unsupported(
      capability,
      "office-js",
      "Excel.run is not available in this runtime",
      "Requires Microsoft Office Excel with Office.js loaded",
    );
  }
  try {
    return ok(
      await run(
        fn as unknown as (ctx: import("./officeJsExcelTypes").ExcelRequestContext) => Promise<T>,
      ),
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return fail(capability, "office-js", message);
  }
}

export function queueLoadSlicer(slicer: ExcelSlicer): void {
  if (typeof slicer.load !== "function") {
    throw new Error("Slicer.load is not available");
  }
  slicer.load(SLICER_SCALAR_PROPS);
  if (!slicer.worksheet || typeof slicer.worksheet.load !== "function") {
    throw new Error("Slicer.worksheet.load is not available");
  }
  slicer.worksheet.load("name");
}

export function assertFiniteNonNegative(label: string, value: number, allowZero: boolean): void {
  if (!Number.isFinite(value)) throw new Error(`${label} must be a finite number`);
  if (allowZero) {
    if (value < 0) throw new Error(`${label} must be >= 0`);
  } else if (value <= 0) {
    throw new Error(`${label} must be > 0`);
  }
}

export function applyWritableSlicerFields(
  slicer: ExcelSlicer,
  fields: {
    name?: string;
    caption?: string;
    top?: number;
    left?: number;
    width?: number;
    height?: number;
    style?: string;
    sortBy?: SlicerSortBy;
  },
): void {
  if (fields.name !== undefined) slicer.name = fields.name;
  if (fields.caption !== undefined) slicer.caption = fields.caption;
  if (fields.top !== undefined) {
    assertFiniteNonNegative("top", fields.top, true);
    slicer.top = fields.top;
  }
  if (fields.left !== undefined) {
    assertFiniteNonNegative("left", fields.left, true);
    slicer.left = fields.left;
  }
  if (fields.width !== undefined) {
    assertFiniteNonNegative("width", fields.width, false);
    slicer.width = fields.width;
  }
  if (fields.height !== undefined) {
    assertFiniteNonNegative("height", fields.height, false);
    slicer.height = fields.height;
  }
  if (fields.style !== undefined) slicer.style = fields.style;
  if (fields.sortBy !== undefined) slicer.sortBy = mapSortByToHost(fields.sortBy);
}
