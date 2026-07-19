import type { HostAdapter } from "../host/types";
import type { ToolCall, ToolResult } from "./types";

const MAX_DIM = 4096;

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

function optionalPositiveInt(args: Record<string, unknown>, key: string): number | undefined {
  if (!Object.prototype.hasOwnProperty.call(args, key)) return undefined;
  if (args[key] === undefined) throw new Error(`${key} must not be undefined`);
  if (args[key] === null) throw new Error(`${key} must not be null`);
  const value = args[key];
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    !Number.isInteger(value) ||
    value < 1 ||
    value > MAX_DIM
  ) {
    throw new Error(`${key} must be an integer from 1 to ${MAX_DIM}`);
  }
  return value;
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

export async function executeChartImageTool(
  host: HostAdapter,
  call: ToolCall,
): Promise<ToolResult | null> {
  if (call.name !== "chart.image.get") return null;
  rejectUnknown(call.arguments, ["sheetName", "chartName", "width", "height"]);
  const width = optionalPositiveInt(call.arguments, "width");
  const height = optionalPositiveInt(call.arguments, "height");
  return fromHost(
    call.name,
    await host.getChartImage({
      sheetName: requireString(call.arguments, "sheetName"),
      chartName: requireString(call.arguments, "chartName"),
      ...(width != null ? { width } : {}),
      ...(height != null ? { height } : {}),
    }),
  );
}
