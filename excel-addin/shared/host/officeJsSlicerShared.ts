/**
 * Shared Slicer runner, sort mapping, and snapshot readback (ExcelApi 1.10).
 */
import { requireExcelApi110ForSlicer } from "./officeJsSlicerRequirements";
import type {
  ExcelRequestContextWithSlicer,
  ExcelSlicer,
  ExcelSlicerSortHost,
} from "./officeJsSlicerTypes";
import type { SlicerInfo, SlicerSortBy } from "./slicerTypes";
import { getExcelRun } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { fail, ok, unsupported } from "./types";

export const SLICER_SCALAR_PROPS =
  "name,id,caption,height,width,top,left,sortBy,style,isFilterCleared";

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

/** Case-insensitive only; no trim/punctuation strip (Phase53.3 lesson). */
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
  // Reject whitespace variants: "Ascending " must fail (toLowerCase alone keeps spaces).
  if (raw.toLowerCase() !== raw.toLowerCase().trim() || /\s/.test(raw)) {
    throw new Error(`unknown slicer sortBy host token: ${raw}`);
  }
  // Official host tokens have no spaces; ensure exact structure after case fold.
  const expected = SORT_TO_HOST[mapped];
  if (raw.toLowerCase() !== expected.toLowerCase()) {
    throw new Error(`unknown slicer sortBy host token: ${raw}`);
  }
  return mapped;
}

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
  if (slicer.worksheet && typeof slicer.worksheet.load === "function") {
    slicer.worksheet.load("name");
  }
}

export function readSlicerSnapshot(slicer: ExcelSlicer, limitations: string[] = []): SlicerInfo {
  const sheetName =
    slicer.worksheet && typeof slicer.worksheet.name === "string"
      ? String(slicer.worksheet.name)
      : "";
  if (!sheetName) limitations.push("worksheet name unavailable after load");

  const sortBy = mapSortByFromHost(slicer.sortBy as ExcelSlicerSortHost);
  const id = slicer.id == null ? "" : String(slicer.id);
  if (!id) limitations.push("slicer id empty after load");

  const info: SlicerInfo = {
    name: String(slicer.name ?? ""),
    id,
    caption: String(slicer.caption ?? ""),
    sheetName,
    top: Number(slicer.top),
    left: Number(slicer.left),
    width: Number(slicer.width),
    height: Number(slicer.height),
    sortBy,
    style: String(slicer.style ?? ""),
    isFilterCleared: Boolean(slicer.isFilterCleared),
  };
  if (limitations.length > 0) info.limitations = [...limitations];
  return info;
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
