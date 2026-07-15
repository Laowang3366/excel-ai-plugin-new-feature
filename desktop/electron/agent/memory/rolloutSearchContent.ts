import type { RolloutItem } from "../shared/types";
import { redactSensitiveText, summarizeValueForAudit } from "../../shared/sensitiveData";

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
      return redactSensitiveText(turnItem.content, 100_000);
    }
    if (turnItem.type === "tool_call") {
      return `${turnItem.toolName} ${summarizeValueForAudit(turnItem.arguments)}`;
    }
    if (turnItem.type === "tool_result") {
      return `${turnItem.toolName} ${turnItem.isError ? "error" : "success"} ${summarizeValueForAudit(turnItem.result)}`;
    }
    if (turnItem.type === "reasoning") {
      return redactSensitiveText(turnItem.summaryText.join(" "), 100_000);
    }
    if (turnItem.type === "compacted") {
      return redactSensitiveText(turnItem.summary, 100_000);
    }
    if (turnItem.type === "compact_progress") {
      return redactSensitiveText(`${turnItem.reason} ${turnItem.status}`, 100_000);
    }
  }

  if (item.type === "session_meta") {
    return redactSensitiveText(
      [item.meta.id, item.meta.modelProvider, item.meta.model, item.meta.folderId]
        .filter(Boolean)
        .join(" "),
      100_000,
    );
  }
  if (item.type === "compacted") return redactSensitiveText(item.summary, 100_000);
  if (item.type === "compact_params") {
    return redactSensitiveText(`${item.reason} ${item.status}`, 100_000);
  }
  if (item.type === "turn_usage") return "turn usage";
  if (item.type === "turn_context") return redactSensitiveText(item.cwd ?? "", 100_000);

  return summarizeValueForAudit(item);
}
