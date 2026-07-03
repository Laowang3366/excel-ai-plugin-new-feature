import type { RolloutItem } from "../shared/types";

/**
 * Rollout 检索文本提取。
 *
 * 关联模块：
 * - stateRuntimeStore.ts: 写入 logs.db FTS 索引时使用。
 * - rolloutArchive.ts: 扫描压缩 JSONL 归档时使用同一套文本口径。
 */
export function extractRolloutSearchContent(item: RolloutItem): string {
  if (item.type === "turn_item") {
    const turnItem = item.item;
    if (turnItem.type === "user_message" || turnItem.type === "assistant_message") {
      return turnItem.content;
    }
    if (turnItem.type === "tool_call") {
      return `${turnItem.toolName} ${JSON.stringify(turnItem.arguments)}`;
    }
    if (turnItem.type === "tool_result") {
      return `${turnItem.toolName} ${stringifyUnknown(turnItem.result)}`;
    }
    if (turnItem.type === "reasoning") {
      return [...turnItem.summaryText, ...turnItem.rawContent].join(" ");
    }
    if (turnItem.type === "compacted") {
      return turnItem.summary;
    }
    if (turnItem.type === "compact_progress") {
      return `${turnItem.reason} ${turnItem.status}`;
    }
  }

  if (item.type === "session_meta") {
    return [
      item.meta.id,
      item.meta.modelProvider,
      item.meta.model,
      item.meta.folderId,
    ].filter(Boolean).join(" ");
  }
  if (item.type === "compacted") return item.summary;
  if (item.type === "compact_params") return `${item.reason} ${item.status}`;
  if (item.type === "turn_usage") return "turn usage";
  if (item.type === "turn_context") return item.cwd ?? "";

  return JSON.stringify(item);
}

function stringifyUnknown(value: unknown): string {
  return typeof value === "string" ? value : JSON.stringify(value);
}
