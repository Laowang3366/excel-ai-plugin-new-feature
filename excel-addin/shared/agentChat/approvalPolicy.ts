import type { RiskLevel } from "../tools/types";

export type ApprovalDisposition = "direct" | "approval" | "deny";

/**
 * Desktop-aligned permission modes (settingsStore / toolApproval.ts):
 * - normal: confirm every tool call (逐次确认)
 * - auto_approve_safe: only safe tools auto-run (自动批准) — add-in default
 * - confirm_all: full access, no approval dialogs (完整权限)
 */
export type PermissionMode = "normal" | "auto_approve_safe" | "confirm_all";

export const PERMISSION_MODES: readonly PermissionMode[] = [
  "normal",
  "auto_approve_safe",
  "confirm_all",
] as const;

/** Keep current add-in safety: safe auto, moderate/dangerous need approval. */
export const DEFAULT_PERMISSION_MODE: PermissionMode = "auto_approve_safe";

export const PERMISSION_MODE_LABELS: Record<PermissionMode, string> = {
  normal: "逐次确认",
  auto_approve_safe: "自动批准安全操作",
  confirm_all: "完整权限（自动执行）",
};

export const CHAT_APPROVAL_REJECT_PREFIX = "chat approval: rejected by user:";
export const CHAT_APPROVAL_DENY_PREFIX = "chat approval: tool not allowed:";

const PERMISSION_MODE_SET = new Set<string>(PERMISSION_MODES);

export function isPermissionMode(value: unknown): value is PermissionMode {
  return typeof value === "string" && PERMISSION_MODE_SET.has(value);
}

/** Invalid / missing values fall back to the safe default. */
export function normalizePermissionMode(value: unknown): PermissionMode {
  return isPermissionMode(value) ? value : DEFAULT_PERMISSION_MODE;
}

function isRiskLevel(value: unknown): value is RiskLevel {
  return value === "safe" || value === "moderate" || value === "dangerous";
}

/**
 * Map risk × permission mode → disposition.
 * Unknown risk always denies (never silent execute).
 * confirm_all auto-runs known tools only — registry miss still denies upstream.
 */
export function dispositionForRisk(
  risk: RiskLevel | string | undefined | null,
  mode: PermissionMode | string | undefined | null = DEFAULT_PERMISSION_MODE,
): ApprovalDisposition {
  if (!isRiskLevel(risk)) return "deny";

  const permissionMode = normalizePermissionMode(mode);

  switch (permissionMode) {
    case "confirm_all":
      return "direct";
    case "normal":
      return "approval";
    case "auto_approve_safe":
    default:
      return risk === "safe" ? "direct" : "approval";
  }
}

export function rejectedToolError(toolName: string): string {
  return `${CHAT_APPROVAL_REJECT_PREFIX} ${toolName}`;
}

export function deniedToolError(toolName: string): string {
  return `${CHAT_APPROVAL_DENY_PREFIX} ${toolName}`;
}
