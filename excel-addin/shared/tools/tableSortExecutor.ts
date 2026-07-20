import type { HostAdapter } from "../host/types";
import type { TableSortFieldInput } from "../host/tableSortTypes";
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

function optionalBoolean(args: Record<string, unknown>, key: string): boolean | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key) || args[key] === undefined) return undefined;
  if (args[key] === null) throw new Error(`${key} must not be null`);
  if (typeof args[key] !== "boolean") throw new Error(`${key} must be a boolean`);
  return args[key] as boolean;
}

function requireFields(args: Record<string, unknown>): TableSortFieldInput[] {
  if (!Object.prototype.hasOwnProperty.call(args, "fields")) {
    throw new Error("Missing fields argument");
  }
  if (args.fields === null) throw new Error("fields must not be null");
  if (!Array.isArray(args.fields) || args.fields.length === 0) {
    throw new Error("fields must be a non-empty array");
  }
  if (args.fields.length > 3) throw new Error("fields supports at most 3 levels");
  return args.fields.map((raw, index) => {
    if (raw == null || typeof raw !== "object" || Array.isArray(raw)) {
      throw new Error(`fields[${index}] must be an object`);
    }
    const field = raw as Record<string, unknown>;
    for (const key of Object.keys(field)) {
      if (key !== "columnIndex" && key !== "ascending") {
        throw new Error(`fields[${index}] unknown field: ${key}`);
      }
    }
    const columnIndex = field.columnIndex;
    if (typeof columnIndex !== "number" || !Number.isInteger(columnIndex) || columnIndex < 1) {
      throw new Error(`fields[${index}].columnIndex must be a 1-based integer >= 1`);
    }
    let ascending: boolean | undefined;
    if (Object.prototype.hasOwnProperty.call(field, "ascending") && field.ascending !== undefined) {
      if (field.ascending === null) throw new Error(`fields[${index}].ascending must not be null`);
      if (typeof field.ascending !== "boolean") {
        throw new Error(`fields[${index}].ascending must be a boolean`);
      }
      ascending = field.ascending;
    }
    return { columnIndex, ascending };
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

export async function executeTableSortTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name === "table.sort.get") {
    rejectUnknown(call.arguments, ["sheetName", "tableName"]);
    return fromHost(
      call.name,
      await host.getTableSort({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
      }),
    );
  }
  if (call.name === "table.sort.apply") {
    rejectUnknown(call.arguments, ["sheetName", "tableName", "fields", "matchCase"]);
    return fromHost(
      call.name,
      await host.applyTableSort({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
        fields: requireFields(call.arguments),
        matchCase: optionalBoolean(call.arguments, "matchCase"),
      }),
    );
  }
  if (call.name === "table.sort.clear") {
    rejectUnknown(call.arguments, ["sheetName", "tableName"]);
    return fromHost(
      call.name,
      await host.clearTableSort({
        sheetName: requireString(call.arguments, "sheetName"),
        tableName: requireString(call.arguments, "tableName"),
      }),
    );
  }
  return null;
}
