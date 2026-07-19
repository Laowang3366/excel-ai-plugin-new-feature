import type { RiskLevel } from "../tools/types";

export type ApprovalDisposition = "direct" | "approval" | "deny";

export const CHAT_APPROVAL_REJECT_PREFIX = "chat approval: rejected by user:";
export const CHAT_APPROVAL_DENY_PREFIX = "chat approval: tool not allowed:";

export function dispositionForRisk(
  risk: RiskLevel | string | undefined | null,
): ApprovalDisposition {
  if (risk === "safe") return "direct";
  if (risk === "moderate" || risk === "dangerous") return "approval";
  return "deny";
}

export function rejectedToolError(toolName: string): string {
  return `${CHAT_APPROVAL_REJECT_PREFIX} ${toolName}`;
}

export function deniedToolError(toolName: string): string {
  return `${CHAT_APPROVAL_DENY_PREFIX} ${toolName}`;
}
