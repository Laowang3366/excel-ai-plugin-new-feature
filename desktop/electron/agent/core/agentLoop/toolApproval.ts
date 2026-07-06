import { TOOL_DEFINITIONS_MAP } from "../../tools/registry/toolDefinitions";

export interface ToolApprovalConfig {
  permissionMode: "normal" | "auto_approve_safe" | "confirm_all";
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    sandboxJustification?: string;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
}

const alwaysAllowedTools = new Set<string>();

export function shouldRequireApproval(
  toolName: string,
  permissionMode: ToolApprovalConfig["permissionMode"] = "normal"
): boolean {
  switch (permissionMode) {
    case "normal":
      if (alwaysAllowedTools.has(toolName)) return false;
      return true;

    case "auto_approve_safe": {
      if (alwaysAllowedTools.has(toolName)) return false;
      const safeDef = TOOL_DEFINITIONS_MAP.get(toolName);
      return safeDef ? safeDef.riskLevel !== "safe" : true;
    }

    case "confirm_all":
      return false;

    default:
      return true;
  }
}

export async function requestToolApproval(
  params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    sandboxJustification?: string;
  },
  config: ToolApprovalConfig
): Promise<{ approved: boolean; alwaysAllow?: boolean }> {
  if (config.requestToolApproval) {
    return config.requestToolApproval(params);
  }
  return { approved: true };
}

export function markToolAlwaysAllowed(toolName: string): void {
  alwaysAllowedTools.add(toolName);
}

export function getAlwaysAllowedTools(): ReadonlySet<string> {
  return alwaysAllowedTools;
}

export function clearAlwaysAllowedTools(): void {
  alwaysAllowedTools.clear();
}
