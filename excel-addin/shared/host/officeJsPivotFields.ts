/**
 * Pivot field parsing, conflict checks, and hierarchy apply/readback.
 */
import type {
  PivotAggregationFunction,
  PivotCreateInput,
  PivotFieldSpec,
  PivotHierarchySummary,
  PivotNormalizedField,
} from "./pivotTypes";
import { PIVOT_AGGREGATION_FUNCTIONS, PIVOT_MAX_FIELDS } from "./pivotTypes";
import type {
  ExcelAggregationFunction,
  ExcelDataPivotHierarchy,
  ExcelPivotHierarchy,
  ExcelPivotTable,
} from "./officeJsPivotTypes";

const AGG_TO_HOST: Record<PivotAggregationFunction, ExcelAggregationFunction> = {
  sum: "Sum",
  count: "Count",
  average: "Average",
  max: "Max",
  min: "Min",
};

const HOST_TO_AGG: Record<string, PivotAggregationFunction> = {
  sum: "sum",
  count: "count",
  average: "average",
  max: "max",
  min: "min",
};

export type PivotFieldPlan = {
  rowFields: PivotNormalizedField[];
  columnFields: PivotNormalizedField[];
  filterFields: PivotNormalizedField[];
  dataFields: PivotNormalizedField[];
};

function normalizeOne(spec: PivotFieldSpec, axis: string, allowFunction: boolean): PivotNormalizedField {
  if (typeof spec === "string") {
    const name = spec.trim();
    if (name === "") throw new Error(`${axis} field name must be non-empty`);
    return { name };
  }
  if (spec == null || typeof spec !== "object" || Array.isArray(spec)) {
    throw new Error(`${axis} field must be a string or {name,function?,caption?}`);
  }
  const bag = spec as Record<string, unknown>;
  for (const key of Object.keys(bag)) {
    if (key !== "name" && key !== "function" && key !== "caption") {
      throw new Error(`unknown field property: ${key}`);
    }
  }
  if (typeof bag.name !== "string" || bag.name.trim() === "") {
    throw new Error(`${axis} field name must be non-empty`);
  }
  const name = bag.name.trim();
  let fn: PivotAggregationFunction | undefined;
  if (Object.prototype.hasOwnProperty.call(bag, "function")) {
    if (!allowFunction) {
      throw new Error(`${axis} does not accept function (only dataFields)`);
    }
    if (typeof bag.function !== "string") throw new Error("dataFields function must be a string");
    const raw = bag.function.trim().toLowerCase();
    if (!(PIVOT_AGGREGATION_FUNCTIONS as readonly string[]).includes(raw)) {
      throw new Error("dataFields function must be sum|count|average|max|min");
    }
    fn = raw as PivotAggregationFunction;
  }
  let caption: string | undefined;
  if (Object.prototype.hasOwnProperty.call(bag, "caption")) {
    if (typeof bag.caption !== "string") throw new Error("caption must be a string");
    caption = bag.caption;
  }
  return { name, function: fn, caption };
}

function normalizeAxis(
  specs: PivotFieldSpec[] | undefined,
  axis: string,
  allowFunction: boolean,
): PivotNormalizedField[] {
  if (specs == null) return [];
  if (!Array.isArray(specs)) throw new Error(`${axis} must be an array`);
  if (specs.length > PIVOT_MAX_FIELDS) {
    throw new Error(`${axis} supports at most ${PIVOT_MAX_FIELDS} fields`);
  }
  const out: PivotNormalizedField[] = [];
  const seen = new Set<string>();
  for (const spec of specs) {
    const field = normalizeOne(spec, axis, allowFunction);
    const key = field.name.toLowerCase();
    if (seen.has(key)) throw new Error(`duplicate field in ${axis}: ${field.name}`);
    seen.add(key);
    out.push(field);
  }
  return out;
}

/** Parse and validate field arrays; reject axis conflicts (row/column/filter). */
export function buildPivotFieldPlan(input: PivotCreateInput): PivotFieldPlan {
  const rowFields = normalizeAxis(input.rowFields, "rowFields", false);
  const columnFields = normalizeAxis(input.columnFields, "columnFields", false);
  const filterFields = normalizeAxis(input.filterFields, "filterFields", false);
  const dataFields = normalizeAxis(input.dataFields, "dataFields", true).map((f) => ({
    ...f,
    function: f.function ?? "sum",
    caption: f.caption ?? `汇总项: ${f.name}`,
  }));

  const orientation = new Map<string, string>();
  for (const [axis, fields] of [
    ["rowFields", rowFields],
    ["columnFields", columnFields],
    ["filterFields", filterFields],
  ] as const) {
    for (const f of fields) {
      const key = f.name.toLowerCase();
      const prior = orientation.get(key);
      if (prior) {
        throw new Error(`field "${f.name}" cannot appear in both ${prior} and ${axis}`);
      }
      orientation.set(key, axis);
    }
  }
  return { rowFields, columnFields, filterFields, dataFields };
}

function getHierarchy(pivot: ExcelPivotTable, name: string): ExcelPivotHierarchy {
  const h = pivot.hierarchies.getItem(name);
  if (!h) throw new Error(`pivot hierarchy not found: ${name}`);
  return h;
}

/** Queue hierarchy layout mutations (caller syncs). */
export function applyPivotFieldPlan(pivot: ExcelPivotTable, plan: PivotFieldPlan): void {
  for (const f of plan.rowFields) {
    pivot.rowHierarchies.add(getHierarchy(pivot, f.name));
  }
  for (const f of plan.columnFields) {
    pivot.columnHierarchies.add(getHierarchy(pivot, f.name));
  }
  for (const f of plan.filterFields) {
    pivot.filterHierarchies.add(getHierarchy(pivot, f.name));
  }
  for (const f of plan.dataFields) {
    const data = pivot.dataHierarchies.add(getHierarchy(pivot, f.name)) as ExcelDataPivotHierarchy;
    data.summarizeBy = AGG_TO_HOST[f.function ?? "sum"];
    if (f.caption != null) data.name = f.caption;
  }
}

export function queueLoadPivotHierarchies(pivot: ExcelPivotTable): void {
  pivot.load("name");
  pivot.rowHierarchies.load("items/name");
  pivot.columnHierarchies.load("items/name");
  pivot.filterHierarchies.load("items/name");
  pivot.dataHierarchies.load("items/name,summarizeBy");
  if (pivot.worksheet) pivot.worksheet.load("name");
  const layoutRange = pivot.layout.getRange();
  layoutRange.load("address");
}

export function readHierarchyNames(
  collection: { items?: Array<{ name?: string }> } | undefined,
): PivotHierarchySummary[] {
  const items = collection?.items ?? [];
  return items
    .map((item) => ({ name: String(item.name ?? "") }))
    .filter((item) => item.name !== "")
    .sort((a, b) => a.name.localeCompare(b.name));
}

export function readDataHierarchySummaries(
  collection: { items?: Array<{ name?: string; summarizeBy?: string }> } | undefined,
): PivotHierarchySummary[] {
  const items = collection?.items ?? [];
  return items
    .map((item) => {
      const caption = String(item.name ?? "");
      const host = String(item.summarizeBy ?? "").toLowerCase();
      const summarizeBy = HOST_TO_AGG[host] ?? (item.summarizeBy ? String(item.summarizeBy) : undefined);
      return { name: caption, caption, summarizeBy };
    })
    .filter((item) => item.name !== "")
    .sort((a, b) => a.name.localeCompare(b.name));
}

export function mapAggregationToHost(fn: PivotAggregationFunction): ExcelAggregationFunction {
  return AGG_TO_HOST[fn];
}
