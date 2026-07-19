import type { HostAdapter } from "../host/types";
import type { ToolResult } from "./types";

export async function executeSheetOperation(
  host: HostAdapter,
  args: Record<string, unknown>,
  fromHost: (
    tool: "sheet.operation",
    result: {
      ok: boolean;
      data?: unknown;
      reason?: string;
      unsupported?: boolean;
    },
  ) => ToolResult,
  helpers: {
    requireString: (a: Record<string, unknown>, k: string) => string;
    optionalString: (a: Record<string, unknown>, k: string) => string | undefined;
    optionalFiniteNumber: (a: Record<string, unknown>, k: string) => number | undefined;
    requireSheetOperation: (
      a: Record<string, unknown>,
    ) => "add" | "rename" | "delete" | "copy" | "move";
  },
): Promise<ToolResult> {
  const operation = helpers.requireSheetOperation(args);
  const sheetName = helpers.requireString(args, "sheetName");
  if (operation === "add") {
    return fromHost("sheet.operation", await host.addSheet(sheetName));
  }
  if (operation === "rename") {
    return fromHost(
      "sheet.operation",
      await host.renameSheet(sheetName, helpers.requireString(args, "newName")),
    );
  }
  if (operation === "delete") {
    return fromHost("sheet.operation", await host.deleteSheet(sheetName));
  }
  if (operation === "copy") {
    return fromHost(
      "sheet.operation",
      await host.copySheet(sheetName, helpers.optionalString(args, "newName")),
    );
  }
  const position = helpers.optionalFiniteNumber(args, "position");
  if (position === undefined) {
    throw new Error("position is required for move");
  }
  if (!Number.isInteger(position) || position < 1) {
    throw new Error("position must be a 1-based positive integer");
  }
  return fromHost("sheet.operation", await host.moveSheet(sheetName, position));
}
