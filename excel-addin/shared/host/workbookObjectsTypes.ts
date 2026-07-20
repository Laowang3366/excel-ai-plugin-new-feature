import type { ChartInfo } from "./chartTypes";
import type { ShapeInfo } from "./shapeTypes";
import type { NamedRangeInfo, SheetInfo, TableInfo } from "./types";

/** Per-category inventory status for workbook.objects.inspect. */
export type ObjectCategoryStatus = "available" | "unsupported" | "failed";

/**
 * One object family in the inventory.
 * - available: items may be truncated; totalCount is always the true pre-truncate size
 * - unsupported: items empty; reason/evidence required
 * - failed: items empty; reason required; totalCount null when unknown
 */
export interface ObjectCategoryResult<T> {
  status: ObjectCategoryStatus;
  /** True pre-truncate count when known; null when unknown (failed before enumeration). */
  totalCount: number | null;
  items: T[];
  truncated: boolean;
  reason?: string;
  evidence?: string;
  limitations?: string[];
}

export interface WorkbookObjectsInspectInput {
  /** Default 100; closed range 1..500. */
  maxItemsPerCategory?: number;
  /** When set, only objects on this sheet (named ranges: workbook scope always included). */
  sheetName?: string;
}

export interface WorkbookObjectsInspectInfo {
  workbookName: string;
  activeSheetName: string;
  sheetCount: number;
  /** Sheets included after optional filter; sorted by index then name. */
  sheets: SheetInfo[];
  tables: ObjectCategoryResult<TableInfo>;
  charts: ObjectCategoryResult<ChartInfo>;
  namedRanges: ObjectCategoryResult<NamedRangeInfo>;
  shapes: ObjectCategoryResult<ShapeInfo>;
  /** Top-level limitations (host gaps, truncation policy, inventory slimming). */
  limitations: string[];
  /** Echo of filter when applied. */
  filterSheetName?: string;
}

export const WORKBOOK_OBJECTS_MAX_DEFAULT = 100;
export const WORKBOOK_OBJECTS_MAX_MIN = 1;
export const WORKBOOK_OBJECTS_MAX_MAX = 500;
