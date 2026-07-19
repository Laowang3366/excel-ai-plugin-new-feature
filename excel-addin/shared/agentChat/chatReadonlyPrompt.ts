import {
  composeExcelSystemPrompt,
  type ExcelPromptBuildOptions,
} from "../prompts/composeExcelPrompt";
import type { ToolDefinition } from "../tools/types";
import { listChatReadOnlyTools } from "./chatReadOnlyTools";

/** Stable marker for tests / UI to assert the chat readonly boundary is present. */
export const CHAT_READONLY_PROMPT_MARKER = "CHAT_READONLY_BOUNDARY";

export type ChatReadonlyPromptOptions = ExcelPromptBuildOptions & {
  /** Defaults to listChatReadOnlyTools(). */
  tools?: ToolDefinition[];
};

/**
 * Compose the Excel system prompt, then append a short final chat-only boundary.
 * Does not edit generated desktop templates or advancedExcelBoundary.
 */
export function composeChatReadonlySystemPrompt(
  options: ChatReadonlyPromptOptions,
): string {
  const base = composeExcelSystemPrompt(options);
  const tools = options.tools ?? listChatReadOnlyTools();
  const names = tools.map((t) => t.name).join(", ");
  const boundary = [
    `## ${CHAT_READONLY_PROMPT_MARKER}`,
    "当前聊天为只读模式：写/删/改/保护/宏等变更操作不可用。",
    "仅允许读取与检查类工具；不得调用 set/write/delete/create/update 等写操作。",
    `可用工具：${names || "(none)"}。`,
  ].join("\n");
  return `${base}\n\n${boundary}\n`;
}
