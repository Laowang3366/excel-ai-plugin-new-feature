import type { HostAdapter } from "../host/hostAdapter";
import type {
  SlicerCreateInput,
  SlicerDeleteInput,
  SlicerFilterApplyInput,
  SlicerFilterClearInput,
  SlicerFilterGetInput,
  SlicerListInput,
  SlicerSortBy,
  SlicerSourceType,
  SlicerUpdateInput,
} from "../host/slicerTypes";
import {
  SLICER_MAX_CAPTION_LEN,
  SLICER_MAX_FILTER_KEYS,
  SLICER_MAX_NAME_LEN,
  SLICER_MAX_STYLE_LEN,
  SLICER_SORT_BY_VALUES,
} from "../host/slicerTypes";
import type { ToolCall, ToolResult } from "./types";
import { mapHostResultToToolResult } from "./hostResultMapping";

const LIST_KEYS = new Set(["sheetName"]);
const CREATE_KEYS = new Set([
  "advancedIntent",
  "sourceType",
  "sourceName",
  "sourceField",
  "destinationSheet",
  "name",
  "caption",
  "top",
  "left",
  "width",
  "height",
  "style",
  "sortBy",
]);
const UPDATE_KEYS = new Set([
  "name",
  "newName",
  "caption",
  "top",
  "left",
  "width",
  "height",
  "style",
  "sortBy",
]);
const NAME_KEYS = new Set(["name"]);
const APPLY_KEYS = new Set(["name", "keys"]);

function rejectUnknown(args: Record<string, unknown>, allowed: Set<string>): void {
  for (const key of Object.keys(args)) {
    if (!allowed.has(key)) throw new Error(`unknown field: ${key}`);
  }
}

function requireNonEmptyString(args: Record<string, unknown>, key: string, max: number): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`missing required field: ${key}`);
  }
  const value = args[key];
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "string") throw new Error(`${key} must be a string`);
  if (value.trim() === "") throw new Error(`${key} must be non-empty`);
  if (value.length > max) throw new Error(`${key} exceeds max length ${max}`);
  return value;
}

function optionalNonEmptyString(
  args: Record<string, unknown>,
  key: string,
  max: number,
): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  const value = args[key];
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "string") throw new Error(`${key} must be a string`);
  if (value.trim() === "") throw new Error(`${key} must be non-empty`);
  if (value.length > max) throw new Error(`${key} exceeds max length ${max}`);
  return value;
}

/** Caption may be empty string (official allows). */
function optionalCaption(args: Record<string, unknown>): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "caption")) return undefined;
  const value = args.caption;
  if (value === null) throw new Error("caption must not be null");
  if (typeof value !== "string") throw new Error("caption must be a string");
  if (value.length > SLICER_MAX_CAPTION_LEN) {
    throw new Error(`caption exceeds max length ${SLICER_MAX_CAPTION_LEN}`);
  }
  return value;
}

function optionalNumber(
  args: Record<string, unknown>,
  key: string,
  opts: { min?: number; exclusiveMin?: number },
): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  const value = args[key];
  if (value === null) throw new Error(`${key} must not be null`);
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw new Error(`${key} must be a finite number`);
  }
  if (opts.exclusiveMin !== undefined && value <= opts.exclusiveMin) {
    throw new Error(`${key} must be > ${opts.exclusiveMin}`);
  }
  if (opts.min !== undefined && value < opts.min) {
    throw new Error(`${key} must be >= ${opts.min}`);
  }
  return value;
}

function optionalSortBy(args: Record<string, unknown>): SlicerSortBy | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "sortBy")) return undefined;
  const value = args.sortBy;
  if (value === null) throw new Error("sortBy must not be null");
  if (typeof value !== "string") throw new Error("sortBy must be a string");
  if (!(SLICER_SORT_BY_VALUES as readonly string[]).includes(value)) {
    throw new Error("sortBy must be dataSourceOrder|ascending|descending");
  }
  return value as SlicerSortBy;
}

function parseLayout(args: Record<string, unknown>): {
  caption?: string;
  top?: number;
  left?: number;
  width?: number;
  height?: number;
  style?: string;
  sortBy?: SlicerSortBy;
} {
  const out: {
    caption?: string;
    top?: number;
    left?: number;
    width?: number;
    height?: number;
    style?: string;
    sortBy?: SlicerSortBy;
  } = {};
  const caption = optionalCaption(args);
  if (caption !== undefined) out.caption = caption;
  const top = optionalNumber(args, "top", { min: 0 });
  if (top !== undefined) out.top = top;
  const left = optionalNumber(args, "left", { min: 0 });
  if (left !== undefined) out.left = left;
  const width = optionalNumber(args, "width", { exclusiveMin: 0 });
  if (width !== undefined) out.width = width;
  const height = optionalNumber(args, "height", { exclusiveMin: 0 });
  if (height !== undefined) out.height = height;
  const style = optionalNonEmptyString(args, "style", SLICER_MAX_STYLE_LEN);
  if (style !== undefined) out.style = style;
  const sortBy = optionalSortBy(args);
  if (sortBy !== undefined) out.sortBy = sortBy;
  return out;
}

export async function executeSlicerTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  switch (call.name) {
    case "slicer.list": {
      rejectUnknown(call.arguments, LIST_KEYS);
      const input: SlicerListInput = {};
      const sheetName = optionalNonEmptyString(call.arguments, "sheetName", SLICER_MAX_NAME_LEN);
      if (sheetName) input.sheetName = sheetName;
      return mapHostResultToToolResult(call.name, await host.listSlicers(input));
    }
    case "slicer.create": {
      rejectUnknown(call.arguments, CREATE_KEYS);
      if (call.arguments.advancedIntent !== "interactive-pivot") {
        throw new Error("advancedIntent must be interactive-pivot");
      }
      const sourceType = call.arguments.sourceType;
      if (sourceType !== "table" && sourceType !== "pivotTable") {
        throw new Error("sourceType must be table or pivotTable");
      }
      const input: SlicerCreateInput = {
        advancedIntent: "interactive-pivot",
        sourceType: sourceType as SlicerSourceType,
        sourceName: requireNonEmptyString(call.arguments, "sourceName", SLICER_MAX_NAME_LEN),
        sourceField: requireNonEmptyString(call.arguments, "sourceField", SLICER_MAX_NAME_LEN),
        destinationSheet: requireNonEmptyString(
          call.arguments,
          "destinationSheet",
          SLICER_MAX_NAME_LEN,
        ),
        ...parseLayout(call.arguments),
      };
      const name = optionalNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN);
      if (name) input.name = name;
      return mapHostResultToToolResult(call.name, await host.createSlicer(input));
    }
    case "slicer.update": {
      rejectUnknown(call.arguments, UPDATE_KEYS);
      const input: SlicerUpdateInput = {
        name: requireNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN),
        ...parseLayout(call.arguments),
      };
      const newName = optionalNonEmptyString(call.arguments, "newName", SLICER_MAX_NAME_LEN);
      if (newName) input.newName = newName;
      const keys = Object.keys(input).filter((k) => k !== "name");
      if (keys.length === 0) throw new Error("empty update: provide at least one writable field");
      return mapHostResultToToolResult(call.name, await host.updateSlicer(input));
    }
    case "slicer.delete": {
      rejectUnknown(call.arguments, NAME_KEYS);
      const input: SlicerDeleteInput = {
        name: requireNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN),
      };
      return mapHostResultToToolResult(call.name, await host.deleteSlicer(input));
    }
    case "slicer.filter.get": {
      rejectUnknown(call.arguments, NAME_KEYS);
      const input: SlicerFilterGetInput = {
        name: requireNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN),
      };
      return mapHostResultToToolResult(call.name, await host.getSlicerFilter(input));
    }
    case "slicer.filter.apply": {
      rejectUnknown(call.arguments, APPLY_KEYS);
      if (!Object.prototype.hasOwnProperty.call(call.arguments, "keys")) {
        throw new Error("missing required field: keys");
      }
      if (!Array.isArray(call.arguments.keys)) throw new Error("keys must be an array");
      if (call.arguments.keys.length > SLICER_MAX_FILTER_KEYS) {
        throw new Error(`keys supports at most ${SLICER_MAX_FILTER_KEYS} items`);
      }
      const seen = new Set<string>();
      const keys: string[] = [];
      for (let i = 0; i < call.arguments.keys.length; i += 1) {
        const item = call.arguments.keys[i];
        if (typeof item !== "string") throw new Error(`keys[${i}] must be a string`);
        if (item.trim() === "") throw new Error(`keys[${i}] must be non-empty`);
        if (item.length > SLICER_MAX_NAME_LEN) {
          throw new Error(`keys[${i}] exceeds max length ${SLICER_MAX_NAME_LEN}`);
        }
        if (seen.has(item)) throw new Error(`keys contains duplicate: ${item}`);
        seen.add(item);
        keys.push(item);
      }
      const input: SlicerFilterApplyInput = {
        name: requireNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN),
        keys,
      };
      return mapHostResultToToolResult(call.name, await host.applySlicerFilter(input));
    }
    case "slicer.filter.clear": {
      rejectUnknown(call.arguments, NAME_KEYS);
      const input: SlicerFilterClearInput = {
        name: requireNonEmptyString(call.arguments, "name", SLICER_MAX_NAME_LEN),
      };
      return mapHostResultToToolResult(call.name, await host.clearSlicerFilter(input));
    }
    default:
      return null;
  }
}
