import {
  composeExcelSystemPrompt,
  type ExcelPromptBuildOptions,
} from "../prompts/composeExcelPrompt";
import type { ToolDefinition } from "../tools/types";
import {
  DEFAULT_PERMISSION_MODE,
  normalizePermissionMode,
  type PermissionMode,
} from "./approvalPolicy";
import { listChatTools } from "./chatToolPolicy";

export const CHAT_APPROVAL_PROMPT_MARKER = "CHAT_APPROVAL_BOUNDARY";

export type ChatApprovalPromptOptions = ExcelPromptBuildOptions & {
  /**
   * Current runtime permission mode (same source as ApprovingToolExecutor).
   * Omitted / invalid → DEFAULT_PERMISSION_MODE (auto_approve_safe).
   */
  permissionMode?: PermissionMode | string | null;
  /** Defaults to listChatTools() — test seam for registry-size independence. */
  tools?: ToolDefinition[];
};

function joinNames(names: string[]): string {
  return names.length > 0 ? names.join(", ") : "(none)";
}

function groupByRisk(tools: ToolDefinition[]): {
  safe: string[];
  approval: string[];
  all: string[];
} {
  const safe: string[] = [];
  const approval: string[] = [];
  const all: string[] = [];
  for (const tool of tools) {
    all.push(tool.name);
    if (tool.riskLevel === "safe") safe.push(tool.name);
    else if (tool.riskLevel === "moderate" || tool.riskLevel === "dangerous") {
      approval.push(tool.name);
    }
  }
  return { safe, approval, all };
}

/**
 * Mode-accurate approval boundary lines (no hardcoded tool counts).
 * Semantics must match dispositionForRisk × ApprovingToolExecutor:
 * - normal: every known tool needs user confirmation (including safe)
 * - auto_approve_safe: safe direct; moderate/dangerous need confirmation
 * - confirm_all: known registered tools auto-run; unknown/unregistered still denied
 *   (never claim the user pre-approved each call)
 */
export function buildChatApprovalBoundaryLines(
  modeInput: PermissionMode | string | null | undefined,
  tools: ToolDefinition[],
): string[] {
  const mode = normalizePermissionMode(modeInput);
  const { safe, approval, all } = groupByRisk(tools);

  const commonTail = [
    "在用户批准前，不得声称已经写入、删除或修改工作簿。",
    "若用户拒绝某次工具调用，应根据失败结果调整方案，不要重复同一危险操作轰炸。",
    "WPS 上部分能力可能为 unsupported；宏/COM/.NET/桌面自动化等硬边界仍然适用。",
    "未在工具注册表中的未知工具一律拒绝，不会静默执行。",
  ];

  if (mode === "normal") {
    return [
      `## ${CHAT_APPROVAL_PROMPT_MARKER}`,
      "当前审批模式：normal（逐次确认）。",
      "所有工具调用（含 safe 只读）均需用户确认后才会执行；在确认前不要假定操作已生效。",
      ...commonTail,
      `需确认后执行的工具：${joinNames(all)}。`,
      `其中 safe：${joinNames(safe)}；moderate/dangerous：${joinNames(approval)}。`,
    ];
  }

  if (mode === "confirm_all") {
    return [
      `## ${CHAT_APPROVAL_PROMPT_MARKER}`,
      "当前审批模式：confirm_all（完整权限 / 自动执行已知工具）。",
      "已注册的已知工具将自动执行，不会弹出逐条批准对话框；这不等于用户已对每次调用做过逐项批准。",
      "未知或未注册工具仍会被拒绝，不得猜测或伪造执行结果。",
      ...commonTail.slice(1), // keep reject/WPS/unknown lines; skip "在用户批准前" which conflicts
      "自动执行仍须遵守宿主能力与硬边界；失败时据实说明，不得伪造成功。",
      `可自动执行的已知工具：${joinNames(all)}。`,
      `其中 safe：${joinNames(safe)}；moderate/dangerous：${joinNames(approval)}。`,
    ];
  }

  // auto_approve_safe (default)
  return [
    `## ${CHAT_APPROVAL_PROMPT_MARKER}`,
    "当前审批模式：auto_approve_safe（自动批准安全操作）。",
    "safe 工具可直接执行；moderate/dangerous 写/删/改/保护等变更操作需用户确认后才会执行。",
    ...commonTail,
    `可直接执行（safe）：${joinNames(safe)}。`,
    `需确认后执行（moderate/dangerous）：${joinNames(approval)}。`,
  ];
}

/**
 * Full-capability chat prompt with an approval boundary as the final section.
 * Does not edit generated desktop templates.
 */
export function composeChatApprovalSystemPrompt(
  options: ChatApprovalPromptOptions,
): string {
  const base = composeExcelSystemPrompt(options);
  const tools = options.tools ?? listChatTools();
  const mode = options.permissionMode ?? DEFAULT_PERMISSION_MODE;
  const boundary = buildChatApprovalBoundaryLines(mode, tools).join("\n");
  return `${base}\n\n${boundary}\n`;
}
