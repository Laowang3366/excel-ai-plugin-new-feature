import type { HostAdapter } from "../host/hostAdapter";
import type {
  PivotCreateInput,
  PivotFieldSpec,
  PivotListInput,
  PivotRefreshInput,
} from "../host/pivotTypes";
import { mapHostResultToToolResult } from "./hostResultMapping";
import type { ToolCall, ToolResult } from "./types";

const CREATE_KEYS = [
  "advancedIntent",
  "sourceSheetName",
  "sourceAddress",
  "name",
  "destination",
  "rowFields",
  "columnFields",
  "filterFields",
  "dataFields",
] as const;

const REFRESH_KEYS = ["advancedIntent", "sheetName", "name", "refreshConnections"] as const;
const LIST_KEYS = ["sheetName"] as const;

function rejectUnknown(args: Record<string, unknown>, allowed: readonly string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function requireIntent(args: Record<string, unknown>): "interactive-pivot" {
  const value = args.advancedIntent;
  if (value !== "interactive-pivot") {
    throw new Error('advancedIntent must be "interactive-pivot"');
  }
  return "interactive-pivot";
}

function requireNonEmptyString(args: Record<string, unknown>, key: string): string {
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

/**
 * Optional non-empty string: property missing → omit.
 * Explicit null / non-string / "" / whitespace → error (never silent omit).
 */
function optionalNonEmptyString(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    return undefined;
  }
  const value = args[key];
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "string") throw new Error(`Invalid string argument: ${key}`);
  const trimmed = value.trim();
  if (trimmed === "") throw new Error(`${key} must be non-empty`);
  return trimmed;
}

/**
 * destination: property missing → omit (default Pivots auto placement).
 * Explicit "" → omit (documented default Pivots semantics).
 * Explicit null / non-string / whitespace-only → error (not silent default).
 */
function optionalDestination(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "destination")) {
    return undefined;
  }
  const value = args.destination;
  if (value === null) throw new Error("destination must not be null");
  if (value === "") return undefined;
  if (typeof value !== "string") {
    throw new Error("Invalid string argument: destination");
  }
  const trimmed = value.trim();
  if (trimmed === "") throw new Error("destination must be non-empty when provided");
  return trimmed;
}

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  const value = args[key];
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "boolean") throw new Error(`Invalid boolean argument: ${key}`);
  return value;
}

/**
 * Parse field specs. function/caption only legal on dataFields (enforced here before Host).
 */
function parseFieldSpecs(value: unknown, axis: string): PivotFieldSpec[] | undefined {
  if (value == null) return undefined;
  if (!Array.isArray(value)) throw new Error(`${axis} must be an array`);
  if (value.length > 64) throw new Error(`${axis} supports at most 64 fields`);
  const allowFunctionCaption = axis === "dataFields";
  return value.map((item, index) => {
    if (typeof item === "string") {
      if (item.trim() === "") throw new Error(`${axis}[${index}] must be non-empty`);
      return item;
    }
    if (item == null || typeof item !== "object" || Array.isArray(item)) {
      throw new Error(`${axis}[${index}] must be a string or object`);
    }
    const bag = item as Record<string, unknown>;
    for (const key of Object.keys(bag)) {
      if (key !== "name" && key !== "function" && key !== "caption") {
        throw new Error(`unknown field property: ${key}`);
      }
    }
    if (typeof bag.name !== "string" || bag.name.trim() === "") {
      throw new Error(`${axis}[${index}].name must be non-empty`);
    }
    const hasFunction = Object.prototype.hasOwnProperty.call(bag, "function");
    const hasCaption = Object.prototype.hasOwnProperty.call(bag, "caption");
    if (!allowFunctionCaption && (hasFunction || hasCaption)) {
      throw new Error(`${axis} does not accept function/caption (only dataFields)`);
    }
    const out: PivotFieldSpec = { name: bag.name.trim() };
    if (hasFunction) {
      if (typeof bag.function !== "string") throw new Error("function must be a string");
      const fn = bag.function.trim().toLowerCase();
      if (!["sum", "count", "average", "max", "min"].includes(fn)) {
        throw new Error("function must be sum|count|average|max|min");
      }
      (out as { function?: string }).function = fn;
    }
    if (hasCaption) {
      if (typeof bag.caption !== "string") throw new Error("caption must be a string");
      (out as { caption?: string }).caption = bag.caption;
    }
    return out;
  });
}

export async function executePivotTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  switch (call.name) {
    case "pivot.list": {
      rejectUnknown(call.arguments, LIST_KEYS);
      const input: PivotListInput = {};
      const sheetName = optionalNonEmptyString(call.arguments, "sheetName");
      if (sheetName) input.sheetName = sheetName;
      return mapHostResultToToolResult(call.name, await host.listPivots(input));
    }
    case "pivot.create": {
      rejectUnknown(call.arguments, CREATE_KEYS);
      const input: PivotCreateInput = {
        advancedIntent: requireIntent(call.arguments),
        sourceSheetName: requireNonEmptyString(call.arguments, "sourceSheetName"),
        sourceAddress: requireNonEmptyString(call.arguments, "sourceAddress"),
      };
      const name = optionalNonEmptyString(call.arguments, "name");
      if (name) input.name = name;
      const destination = optionalDestination(call.arguments);
      if (destination) input.destination = destination;
      const rowFields = parseFieldSpecs(call.arguments.rowFields, "rowFields");
      if (rowFields) input.rowFields = rowFields;
      const columnFields = parseFieldSpecs(call.arguments.columnFields, "columnFields");
      if (columnFields) input.columnFields = columnFields;
      const filterFields = parseFieldSpecs(call.arguments.filterFields, "filterFields");
      if (filterFields) input.filterFields = filterFields;
      const dataFields = parseFieldSpecs(call.arguments.dataFields, "dataFields");
      if (dataFields) input.dataFields = dataFields;
      return mapHostResultToToolResult(call.name, await host.createPivot(input));
    }
    case "pivot.refresh": {
      rejectUnknown(call.arguments, REFRESH_KEYS);
      const input: PivotRefreshInput = {
        advancedIntent: requireIntent(call.arguments),
      };
      const sheetName = optionalNonEmptyString(call.arguments, "sheetName");
      if (sheetName) input.sheetName = sheetName;
      const name = optionalNonEmptyString(call.arguments, "name");
      if (name) input.name = name;
      const refreshConnections = optionalBoolean(call.arguments, "refreshConnections");
      if (refreshConnections !== undefined) input.refreshConnections = refreshConnections;
      return mapHostResultToToolResult(call.name, await host.refreshPivots(input));
    }
    default:
      return null;
  }
}
