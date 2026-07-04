/**
 * 上下文压缩（Compaction）— 参考 Codex 的 compact.rs
 *
 * 核心功能：
 * 1. 当对话 token 数接近上限时，自动压缩历史记录
 * 2. 压缩策略：保留最近的用户消息 + AI 生成的摘要
 * 3. 中断后恢复时，先压缩再继续，避免上下文溢出
 *
 * Codex 的压缩流程：
 *   1. 收集当前历史中的所有用户消息
 *   2. 用 AI 生成一段摘要（"以下是之前对话的总结..."）
 *   3. 构建新的压缩历史 = [初始上下文] + [最近的用户消息] + [摘要]
 *   4. 替换原有历史
 *
 * 我们的适配：
 *   - 压缩摘要生成通过 core/agentLoop/compactionProvider.ts 分发
 *   - 推理过程完整保留，不隐藏
 *   - 压缩后的摘要作为 CompactedItem 记录到 Rollout
 */

import {
  type TurnItem,
  type UserMessageItem,
  type AssistantMessageItem,
  type ReasoningItem,
  type CompactedItem,
  type TokenUsage,
  type ToolDefinition,
  type CompactionReason,
  type CompactionConfig,
  DEFAULT_COMPACTION_CONFIG,
} from "../shared/types";

interface RequestEstimateMessage {
  role?: string;
  content?: unknown;
  toolCalls?: Array<{
    id?: string;
    function?: {
      name?: string;
      arguments?: string;
    };
  }>;
}

export interface RequestTokenEstimateInput {
  messages?: RequestEstimateMessage[];
  systemPrompt?: string;
  tools?: ToolDefinition[];
}

// ============================================================
// 常量
// ============================================================

/** 摘要前缀（参考 Codex SUMMARY_PREFIX） */
export const SUMMARY_PREFIX = "[对话摘要]";

// ============================================================
// Token 估算
// ============================================================

/** 粗略估算文本的 token 数（中文约 1.5 字/token，英文约 4 字符/token） */
export function estimateTokens(text: string): number {
  // 简单估算：混合中英文取平均
  const charCount = text.length;
  const chineseChars = (text.match(/[\u4e00-\u9fff]/g) || []).length;
  const otherChars = charCount - chineseChars;
  // 中文约 1.5 字/token，英文约 0.25 词/token（4字符≈1token）
  return Math.ceil(chineseChars / 1.5 + otherChars / 4);
}

/** 估算实际请求负载：系统提示词、消息和工具 schema */
export function estimateRequestTokens(input: RequestTokenEstimateInput): number {
  let total = 0;

  if (input.systemPrompt) {
    total += estimateTokens(input.systemPrompt) + 4;
  }

  for (const message of input.messages || []) {
    total += 4;
    total += estimateTokens(message.role || "");
    total += estimateMessageContentTokens(message.content);
    if (message.toolCalls?.length) {
      total += estimateTokens(JSON.stringify(message.toolCalls));
      total += message.toolCalls.length * 20;
    }
  }

  if (input.tools?.length) {
    total += estimateTokens(JSON.stringify(input.tools));
    total += input.tools.length * 10;
  }

  return total;
}

function estimateMessageContentTokens(content: unknown): number {
  if (!content) return 0;
  if (typeof content === "string") return estimateTokens(content);
  if (Array.isArray(content)) {
    return content.reduce((sum, part) => sum + estimateContentPartTokens(part), 0);
  }
  return estimateTokens(JSON.stringify(content));
}

function estimateContentPartTokens(part: unknown): number {
  if (!part) return 0;
  if (typeof part === "string") return estimateTokens(part);
  if (typeof part !== "object") return estimateTokens(String(part));

  const value = part as Record<string, unknown>;
  if (typeof value.text === "string") return estimateTokens(value.text);
  if (typeof value.content === "string") return estimateTokens(value.content);
  if (typeof value.output_text === "string") return estimateTokens(value.output_text);
  if (value.image_url) return estimateTokens(JSON.stringify(value.image_url)) + 85;
  if (value.file) return estimateTokens(JSON.stringify(value.file));
  return estimateTokens(JSON.stringify(value));
}

/** 估算 TurnItem 数组的总 token 数 */
export function estimateItemsTokens(items: TurnItem[]): number {
  let total = 0;
  for (const item of items) {
    switch (item.type) {
      case "user_message":
        total += estimateTokens(item.content);
        break;
      case "assistant_message":
        total += estimateTokens(item.content);
        break;
      case "reasoning":
        total += estimateTokens(item.summaryText.join("\n"));
        total += estimateTokens(item.rawContent.join(""));
        break;
      case "tool_call":
        total += estimateTokens(JSON.stringify(item.arguments));
        total += 20; // 工具名开销
        break;
      case "tool_result":
        total += estimateTokens(JSON.stringify(item.result));
        break;
      case "compacted":
        total += estimateTokens(item.summary);
        break;
      case "error":
        total += estimateTokens(item.message);
        break;
    }
  }
  return total;
}

// ============================================================
// 压缩逻辑
// ============================================================

/** 收集历史中的用户消息（参考 Codex collect_user_messages） */
export function collectUserMessages(items: TurnItem[]): UserMessageItem[] {
  return items.filter(
    (item): item is UserMessageItem =>
      item.type === "user_message" && !item.content.startsWith(SUMMARY_PREFIX)
  );
}

/** 判断是否为摘要消息 */
export function isSummaryMessage(content: string): boolean {
  return content.startsWith(SUMMARY_PREFIX);
}

/** 判断是否需要压缩 */
export function shouldCompact(
  items: TurnItem[],
  config: CompactionConfig = DEFAULT_COMPACTION_CONFIG
): boolean {
  if (!config.enabled) return false;
  const tokens = estimateItemsTokens(items);
  return tokens >= config.autoCompactTokenThreshold;
}

/**
 * 构建压缩后的历史记录（参考 Codex build_compacted_history）
 *
 * 策略：
 * 1. 保留最近的用户消息（在 token 预算内）
 * 2. 添加 AI 生成的摘要
 * 3. 保留所有推理过程（不隐藏）
 */
export function buildCompactedHistory(
  userMessages: UserMessageItem[],
  summaryText: string,
  config: CompactionConfig = DEFAULT_COMPACTION_CONFIG
): TurnItem[] {
  const maxTokens = config.retainedUserMessageMaxTokens;
  const maxItems = config.retainedRecentItemCount ?? Number.POSITIVE_INFINITY;
  const selectedMessages: UserMessageItem[] = [];
  let remaining = maxTokens;
  let retainedCount = 0;

  // 从最新的消息开始保留（参考 Codex 的倒序遍历）
  for (let i = userMessages.length - 1; i >= 0; i--) {
    if (retainedCount >= maxItems) break;
    const msg = userMessages[i];
    const tokens = estimateTokens(msg.content);
    if (tokens <= remaining) {
      selectedMessages.unshift(msg);
      remaining -= tokens;
      retainedCount += 1;
    } else if (remaining > 0) {
      // 截断最后一条放不下的消息
      const truncated = msg.content.slice(0, Math.floor(remaining * 1.5));
      selectedMessages.unshift({
        ...msg,
        content: truncated + "...[已截断]",
      });
      retainedCount += 1;
      break;
    } else {
      break;
    }
  }

  // 构建新历史：[用户消息] + [摘要]
  const newHistory: TurnItem[] = [...selectedMessages];

  // 摘要作为用户消息添加（参考 Codex 将摘要编码为 user message）
  newHistory.push({
    type: "user_message",
    id: `compact-summary-${Date.now()}`,
    content: `${SUMMARY_PREFIX}\n${summaryText}`,
    timestamp: Date.now(),
  });

  return newHistory;
}

/**
 * 执行压缩流程
 *
 * @param items 当前历史记录
 * @param summaryText AI 生成的摘要文本
 * @param reason 压缩原因
 * @param config 压缩配置
 * @returns { compactedItem, newHistory } 压缩条目和新的历史记录
 */
export function performCompaction(
  items: TurnItem[],
  summaryText: string,
  reason: CompactionReason,
  config: CompactionConfig = DEFAULT_COMPACTION_CONFIG
): { compactedItem: CompactedItem; newHistory: TurnItem[] } {
  const tokensBefore = estimateItemsTokens(items);
  const userMessages = collectUserMessages(items);
  const newHistory = buildCompactedHistory(userMessages, summaryText, config);
  const tokensAfter = estimateItemsTokens(newHistory);

  const compactedItem: CompactedItem = {
    type: "compacted",
    id: `compacted-${Date.now()}`,
    summary: summaryText,
    tokensBefore,
    tokensAfter,
    reason,
    timestamp: Date.now(),
  };

  return { compactedItem, newHistory };
}

/**
 * 从历史记录中提取对话内容，用于生成摘要
 *
 * 将 TurnItem[] 转换为 AI 可理解的文本格式
 */
export function historyToCompactPrompt(items: TurnItem[]): string {
  const parts: string[] = [];

  for (const item of items) {
    switch (item.type) {
      case "user_message":
        parts.push(`【用户】${item.content}`);
        break;
      case "assistant_message":
        parts.push(`【助手】${item.content}`);
        break;
      case "reasoning":
        if (item.rawContent.length > 0) {
          parts.push(`【思考过程】${item.rawContent.join("")}`);
        } else if (item.summaryText.length > 0) {
          parts.push(`【思考摘要】${item.summaryText.join("\n")}`);
        }
        break;
      case "tool_call":
        parts.push(`【工具调用】${item.toolName}(${JSON.stringify(item.arguments)})`);
        break;
      case "tool_result":
        const resultStr = typeof item.result === "string"
          ? item.result
          : JSON.stringify(item.result, null, 2);
        parts.push(`【工具结果】${item.isError ? "错误: " : ""}${resultStr.slice(0, 500)}`);
        break;
      case "compacted":
        parts.push(`【之前的对话摘要】${item.summary}`);
        break;
      case "error":
        parts.push(`【错误】${item.message}`);
        break;
    }
  }

  return parts.join("\n\n");
}

/**
 * 构建中断恢复的上下文提示
 *
 * 当对话因 max token 或中断而停止时，生成一段提示帮助 AI 理解当前进度
 */
export function buildResumeContext(items: TurnItem[]): string {
  // 找到最后一条助手消息
  const lastAssistant = [...items].reverse().find(
    (item): item is AssistantMessageItem => item.type === "assistant_message"
  );

  // 找到最后一个工具调用
  const lastToolCall = [...items].reverse().find(
    (item): item is import("../shared/types").ToolCallItem => item.type === "tool_call"
  );

  // 找到最后一个推理
  const lastReasoning = [...items].reverse().find(
    (item): item is import("../shared/types").ReasoningItem => item.type === "reasoning"
  );

  const parts: string[] = ["[中断恢复上下文]"];

  if (lastReasoning) {
    parts.push(`AI 最后的思考：${lastReasoning.rawContent.join("") || lastReasoning.summaryText.join("\n")}`);
  }

  if (lastToolCall) {
    const result = [...items].reverse().find(
      (item): item is import("../shared/types").ToolResultItem => item.type === "tool_result" && item.toolCallId === lastToolCall.id
    );
    parts.push(`最后执行的工具：${lastToolCall.toolName}`);
    if (result) {
      parts.push(`工具返回结果：${JSON.stringify(result.result).slice(0, 300)}`);
    }
  }

  if (lastAssistant) {
    parts.push(`AI 最后的回复：${lastAssistant.content.slice(0, 500)}`);
  }

  parts.push("\n请基于以上上下文继续之前的工作。");

  return parts.join("\n");
}
