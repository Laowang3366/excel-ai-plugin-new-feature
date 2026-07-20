import { TOOL_DEFINITION_MAP } from "../tools";
import type { ToolCall, ToolDefinition, ToolName, ToolResult } from "../tools/types";

/**
 * Explicit chat-mode allowlist. Order is stable and intentional.
 * listChatReadOnlyTools() further intersects with riskLevel==="safe".
 */
export const CHAT_READONLY_TOOL_ALLOWLIST = [
  "host.status",
  "selection.get",
  "range.read",
  "range.format.read",
  "formula.read",
  "formula.context",
  "formula.protection.inspect",
  "formula.dependencies.inspect",
  "formula.backups.inspect",
  "sheet.list",
  "table.list",
  "chart.list",
  "chart.series.list",
  "chart.series.trendlines.list",
  "chart.image.get",
  "range.image.get",
  "workbook.inspect",
  "workbook.objects.inspect",
  "conditionalFormat.list",
  "dataValidation.read",
  "sheet.visibility.get",
  "sheet.protection.get",
  "namedRange.list",
  "sheet.display.get",
  "sheet.freeze.get",
  "sheet.pageLayout.get",
  "shape.list",
  "pivot.list",
] as const;

export type ChatReadOnlyToolName = (typeof CHAT_READONLY_TOOL_ALLOWLIST)[number];

export const CHAT_READONLY_DENY_ERROR = "chat readonly: tool not allowed";

const ALLOWLIST_SET = new Set<string>(CHAT_READONLY_TOOL_ALLOWLIST);

export function isChatReadOnlyToolName(name: string): name is ChatReadOnlyToolName {
  return ALLOWLIST_SET.has(name);
}

/**
 * Allowlist ∩ registry ∩ riskLevel==="safe".
 * Unknown or non-safe names are dropped (deny-by-default).
 * Returns a fresh array in allowlist order; does not mutate TOOL_DEFINITIONS.
 */
export function listChatReadOnlyTools(): ToolDefinition[] {
  const out: ToolDefinition[] = [];
  for (const name of CHAT_READONLY_TOOL_ALLOWLIST) {
    const def = Object.prototype.hasOwnProperty.call(TOOL_DEFINITION_MAP, name)
      ? TOOL_DEFINITION_MAP[name as keyof typeof TOOL_DEFINITION_MAP]
      : undefined;
    if (!def) continue;
    if (def.riskLevel !== "safe") continue;
    out.push(def);
  }
  return out;
}

export type AgentToolExecutor = {
  execute(call: ToolCall): Promise<ToolResult>;
};

/**
 * Second-layer chat guard: only allowlist + safe tools reach the host executor.
 * Rejections return a fixed ToolFailure and never throw / never call the inner executor.
 */
export class GuardedChatExecutor implements AgentToolExecutor {
  constructor(private readonly inner: AgentToolExecutor) {}

  async execute(call: ToolCall): Promise<ToolResult> {
    const name = call.name;
    if (!isChatReadOnlyToolName(name)) {
      return deny(name);
    }
    const def = Object.prototype.hasOwnProperty.call(TOOL_DEFINITION_MAP, name)
      ? TOOL_DEFINITION_MAP[name]
      : undefined;
    if (!def || def.riskLevel !== "safe") {
      return deny(name);
    }
    return this.inner.execute(call);
  }
}

function deny(name: string): ToolResult {
  // Rejected names may not be ToolName; cast only on the failure path.
  return {
    ok: false,
    tool: name as ToolName,
    error: `${CHAT_READONLY_DENY_ERROR}: ${name}`,
  };
}
