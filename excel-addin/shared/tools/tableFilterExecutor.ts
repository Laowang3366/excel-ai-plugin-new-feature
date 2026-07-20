import type { HostAdapter } from "../host/types";
import { isTableFilterOn, type TableFilterOn, type TableFilterOperator } from "../host/tableFilterTypes";
import type { ToolCall, ToolResult } from "./types";

function requireString(args: Record<string, unknown>, key: string): string {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing string argument: ${key}`);
  }
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`Missing string argument: ${key}`);
  }
  return value.trim();
}

function optionalString(args: Record<string, unknown>, key: string): string | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) return undefined;
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (typeof value !== "string" || value.trim() === "") {
    throw new Error(`${key} must be a non-empty string`);
  }
  return value.trim();
}

function requirePositiveInt(args: Record<string, unknown>, key: string): number {
  if (!Object.prototype.hasOwnProperty.call(args, key)) {
    throw new Error(`Missing integer argument: ${key}`);
  }
  const value = args[key];
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
    throw new Error(`${key} must be a 1-based integer >= 1`);
  }
  return value;
}

function requireFilterOn(args: Record<string, unknown>): TableFilterOn {
  if (!Object.prototype.hasOwnProperty.call(args, "filterOn")) {
    throw new Error("Missing string argument: filterOn");
  }
  const value = args.filterOn;
  if (typeof value !== "string" || !isTableFilterOn(value)) {
    throw new Error(
      "filterOn must be values|custom|topItems|bottomItems|topPercent|bottomPercent",
    );
  }
  return value;
}

function optionalOperator(args: Record<string, unknown>): TableFilterOperator | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "operator") || args.operator === undefined) {
    return undefined;
  }
  if (args.operator === null) throw new Error("operator must not be null");
  const value = args.operator;
  if (value !== "and" && value !== "or") {
    throw new Error("operator must be and|or");
  }
  return value;
}

function optionalThreshold(args: Record<string, unknown>): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "threshold") || args.threshold === undefined) {
    return undefined;
  }
  if (args.threshold === null) throw new Error("threshold must not be null");
  const value = args.threshold;
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    throw new Error("threshold must be a positive number");
  }
  return value;
}

function optionalValues(args: Record<string, unknown>): string[] | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, "values") || args.values === undefined) {
    return undefined;
  }
  if (args.values === null) throw new Error("values must not be null");
  if (!Array.isArray(args.values) || args.values.length === 0) {
    throw new Error("values must be a non-empty string array");
  }
  if (args.values.length > 256) throw new Error("values supports at most 256 items");
  return args.values.map((item, index) => {
    if (typeof item !== "string") {
      throw new Error(`values[${index}] must be a string`);
    }
    return item;
  });
}

function rejectUnknown(args: Record<string, unknown>, allowed: string[]): void {
  for (const key of Object.keys(args)) {
    if (!allowed.includes(key)) throw new Error(`unknown field: ${key}`);
  }
}

function fromHost(
  tool: ToolCall["name"],
  result: { ok: boolean; data?: unknown; reason?: string; unsupported?: boolean },
): ToolResult {
  if (result.ok) return { ok: true, tool, data: result.data };
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      error: result.reason ?? "host failed",
      detail: result,
      unsupported: true,
    };
  }
  return { ok: false, tool, error: result.reason ?? "host failed", detail: result };
}

export async function executeTableFilterTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "table.filter.get") {
    rejectUnknown(call.arguments, ["sheetName", "tableName"]);
    return fromHost(
      call.name,
      await host.getTableFilter({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
      }),
    );
  }
  if (call.name === "table.filter.apply") {
    rejectUnknown(call.arguments, [
      "sheetName",
      "tableName",
      "columnIndex",
      "filterOn",
      "values",
      "criterion1",
      "criterion2",
      "operator",
      "threshold",
    ]);
    return fromHost(
      call.name,
      await host.applyTableFilter({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
        columnIndex: requirePositiveInt(call.arguments, "columnIndex"),
        filterOn: requireFilterOn(call.arguments),
        values: optionalValues(call.arguments),
        criterion1: optionalString(call.arguments, "criterion1"),
        criterion2: optionalString(call.arguments, "criterion2"),
        operator: optionalOperator(call.arguments),
        threshold: optionalThreshold(call.arguments),
      }),
    );
  }
  if (call.name === "table.filter.clear") {
    rejectUnknown(call.arguments, ["sheetName", "tableName"]);
    return fromHost(
      call.name,
      await host.clearTableFilter({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
      }),
    );
  }
  return null;
}
