import { TOOL_DEFINITIONS_MAP } from "../../tools/registry/toolDefinitions";

export interface ToolApprovalScope {
  threadId: string;
  arguments: Record<string, unknown>;
}

export interface ToolApprovalConfig {
  permissionMode: "normal" | "auto_approve_safe" | "confirm_all";
  requestToolApproval?: (params: {
    toolCallId: string;
    toolName: string;
    arguments: Record<string, unknown>;
    riskLevel: "safe" | "moderate" | "dangerous";
    description?: string;
    canAlwaysAllow?: boolean;
  }) => Promise<{ approved: boolean; alwaysAllow?: boolean }>;
}

const APPROVAL_GRANT_TTL_MS = 30 * 60 * 1000;
const alwaysAllowedScopes = new Map<string, number>();
const DESTRUCTIVE_OPERATIONS = new Set([
  "clear",
  "delete",
  "drop",
  "overwrite",
  "remove",
  "reset",
  "trash",
]);

function getStringArgument(args: Record<string, unknown>, keys: readonly string[]): string | null {
  for (const key of keys) {
    const value = args[key];
    if (typeof value === "string" && value.trim()) return value.trim();
  }
  return null;
}

function getOperation(args: Record<string, unknown>): string {
  return (
    getStringArgument(args, ["operation", "command", "action", "mode"])?.toLowerCase() ?? "default"
  );
}

function getStableTarget(args: Record<string, unknown>): string | null {
  return getStringArgument(args, ["filePath", "workbookPath", "documentPath", "sourcePath"]);
}

function isMandatoryApproval(toolName: string, args: Record<string, unknown>): boolean {
  const toolDef = TOOL_DEFINITIONS_MAP.get(toolName);
  if (!toolDef) return true;
  const operation = getOperation(args);
  if ([...DESTRUCTIVE_OPERATIONS].some((token) => operation.includes(token))) return true;
  if (toolDef.riskLevel === "dangerous" || toolDef.isDataEgress || toolDef.requiresExplicitApproval)
    return true;
  if (!toolDef.isFileDeletion) return false;
  return operation === "default";
}

function buildApprovalKey(toolName: string, scope: ToolApprovalScope): string | null {
  const target = getStableTarget(scope.arguments);
  if (!target) return null;
  return JSON.stringify([
    scope.threadId,
    toolName,
    getOperation(scope.arguments),
    target.toLowerCase(),
  ]);
}

function hasActiveApproval(toolName: string, scope?: ToolApprovalScope): boolean {
  if (!scope || isMandatoryApproval(toolName, scope.arguments)) return false;
  const key = buildApprovalKey(toolName, scope);
  if (!key) return false;
  const expiresAt = alwaysAllowedScopes.get(key);
  if (!expiresAt) return false;
  if (expiresAt <= Date.now()) {
    alwaysAllowedScopes.delete(key);
    return false;
  }
  return true;
}

export function canAlwaysAllowTool(toolName: string, scope: ToolApprovalScope): boolean {
  return (
    !isMandatoryApproval(toolName, scope.arguments) && buildApprovalKey(toolName, scope) !== null
  );
}

export function shouldRequireApproval(
  toolName: string,
  permissionMode: ToolApprovalConfig["permissionMode"] = "normal",
  scope?: ToolApprovalScope,
): boolean {
  if (permissionMode === "confirm_all") return false;

  const args = scope?.arguments ?? {};
  const toolDef = TOOL_DEFINITIONS_MAP.get(toolName);
  if (!toolDef || isMandatoryApproval(toolName, args)) return true;
  if (hasActiveApproval(toolName, scope)) return false;

  switch (permissionMode) {
    case "normal":
      return true;
    case "auto_approve_safe":
      return toolDef.riskLevel !== "safe" || toolDef.requiresApproval;
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
    canAlwaysAllow?: boolean;
  },
  config: ToolApprovalConfig,
): Promise<{ approved: boolean; alwaysAllow?: boolean }> {
  if (!config.requestToolApproval) return { approved: false };
  return config.requestToolApproval(params);
}

export function markToolAlwaysAllowed(toolName: string, scope?: ToolApprovalScope): boolean {
  if (!scope || !canAlwaysAllowTool(toolName, scope)) return false;
  const key = buildApprovalKey(toolName, scope)!;
  alwaysAllowedScopes.set(key, Date.now() + APPROVAL_GRANT_TTL_MS);
  return true;
}

export function getAlwaysAllowedTools(): ReadonlySet<string> {
  const activeTools = new Set<string>();
  for (const [key, expiresAt] of alwaysAllowedScopes) {
    if (expiresAt <= Date.now()) {
      alwaysAllowedScopes.delete(key);
      continue;
    }
    const parsed = JSON.parse(key) as [string, string];
    activeTools.add(parsed[1]);
  }
  return activeTools;
}

export function clearAlwaysAllowedTools(): void {
  alwaysAllowedScopes.clear();
}
