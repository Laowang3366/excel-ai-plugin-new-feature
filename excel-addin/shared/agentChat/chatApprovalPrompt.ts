import {
  composeExcelSystemPrompt,
  type ExcelPromptBuildOptions,
} from "../prompts/composeExcelPrompt";
import { listChatTools } from "./chatToolPolicy";

export const CHAT_APPROVAL_PROMPT_MARKER = "CHAT_APPROVAL_BOUNDARY";

export type ChatApprovalPromptOptions = ExcelPromptBuildOptions;

/**
 * Full-capability chat prompt with an approval boundary as the final section.
 * Does not edit generated desktop templates.
 */
export function composeChatApprovalSystemPrompt(
  options: ChatApprovalPromptOptions,
): string {
  const base = composeExcelSystemPrompt(options);
  const tools = listChatTools();
  const safe = tools.filter((t) => t.riskLevel === "safe").map((t) => t.name);
  const approval = tools
    .filter((t) => t.riskLevel === "moderate" || t.riskLevel === "dangerous")
    .map((t) => t.name);
  const boundary = [
    `## ${CHAT_APPROVAL_PROMPT_MARKER}`,
    "当前聊天可使用全部 Excel 工具，但写/删/改/保护等变更操作需用户逐条批准后才会执行。",
    "在用户批准前，不得声称已经写入、删除或修改工作簿。",
    "若用户拒绝某次工具调用，应根据失败结果调整方案，不要重复同一危险操作轰炸。",
    "WPS 上部分能力可能为 unsupported；宏/COM/.NET/桌面自动化等硬边界仍然适用。",
    `可直接执行（safe）：${safe.join(", ") || "(none)"}。`,
    `需批准后执行：${approval.join(", ") || "(none)"}。`,
  ].join("\n");
  return `${base}\n\n${boundary}\n`;
}
