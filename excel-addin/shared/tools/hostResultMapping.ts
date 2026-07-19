import type { ToolCall, ToolResult } from "./types";

/** HostResult-like shape accepted from HostAdapter methods. */
export type HostResultLike = {
  ok: boolean;
  data?: unknown;
  reason?: string;
  unsupported?: boolean;
  capability?: string;
  host?: string;
  evidence?: string;
};

/**
 * Map HostResult → ToolResult.
 * Only explicit unsupported:true stays typed unsupported; ordinary fail keeps detail without that flag.
 */
export function mapHostResultToToolResult(
  tool: ToolCall["name"],
  result: HostResultLike,
): ToolResult {
  if (result.ok) {
    return { ok: true, tool, data: result.data };
  }
  if (result.unsupported === true) {
    return {
      ok: false,
      tool,
      unsupported: true,
      error: result.reason ?? "unsupported",
      detail: result,
    };
  }
  return {
    ok: false,
    tool,
    error: result.reason ?? "host failed",
    detail: result,
  };
}
