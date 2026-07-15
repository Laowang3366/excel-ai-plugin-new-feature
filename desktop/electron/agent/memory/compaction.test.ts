/**
 * compaction.ts 单元测试
 *
 * 测试上下文压缩的纯函数逻辑：
 * - estimateTokens: token 估算
 * - shouldCompact: 压缩判断
 * - buildCompactedHistory: 压缩历史构建
 * - performCompaction: 压缩执行
 * - historyToCompactPrompt: 历史转 prompt
 * - buildResumeContext: 恢复上下文构建
 */

import { describe, it, expect } from "vitest";
import {
  estimateTokens,
  estimateItemsTokens,
  estimateRequestTokens,
  shouldCompact,
  collectUserMessages,
  buildCompactedHistory,
  performCompaction,
  historyToCompactPrompt,
  buildResumeContext,
  SUMMARY_PREFIX,
} from "./compaction";
import { DEFAULT_COMPACTION_CONFIG, type TurnItem, type CompactionConfig } from "../shared/types";

// ============================================================
// estimateTokens
// ============================================================

describe("estimateTokens", () => {
  it("should estimate English text (~4 chars/token)", () => {
    const tokens = estimateTokens("hello world");
    // 11 chars / 4 ≈ 3
    expect(tokens).toBeGreaterThanOrEqual(2);
    expect(tokens).toBeLessThanOrEqual(5);
  });

  it("should estimate Chinese text (~1.5 chars/token)", () => {
    const tokens = estimateTokens("你好世界测试文本");
    // 8 Chinese chars / 1.5 ≈ 6
    expect(tokens).toBeGreaterThanOrEqual(4);
    expect(tokens).toBeLessThanOrEqual(8);
  });

  it("should handle empty string", () => {
    expect(estimateTokens("")).toBe(0);
  });

  it("should handle mixed content", () => {
    const tokens = estimateTokens("你好world测试");
    expect(tokens).toBeGreaterThan(0);
  });
});

// ============================================================
// estimateItemsTokens
// ============================================================

describe("estimateItemsTokens", () => {
  it("should return 0 for empty array", () => {
    expect(estimateItemsTokens([])).toBe(0);
  });

  it("should estimate user_message tokens", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello world", timestamp: Date.now() },
    ];
    expect(estimateItemsTokens(items)).toBeGreaterThan(0);
  });

  it("should estimate tool_call tokens", () => {
    const items: TurnItem[] = [
      {
        type: "tool_call",
        id: "1",
        toolName: "range.read",
        arguments: { range: "A1" },
        status: "completed",
        timestamp: Date.now(),
      },
    ];
    expect(estimateItemsTokens(items)).toBeGreaterThan(20); // 20 base + args
  });

  it("should sum multiple items", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
      {
        type: "assistant_message",
        id: "2",
        content: "world",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    const single1: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
    ];
    const single2: TurnItem[] = [
      {
        type: "assistant_message",
        id: "2",
        content: "world",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    expect(estimateItemsTokens(items)).toBe(
      estimateItemsTokens(single1) + estimateItemsTokens(single2),
    );
  });
});

describe("estimateRequestTokens", () => {
  it("includes hidden system prompt and tool schema overhead", () => {
    const messageOnly = estimateRequestTokens({
      messages: [{ role: "user", content: "hello" }],
    });
    const withPromptAndTools = estimateRequestTokens({
      systemPrompt: "You are an Office automation agent. Always inspect connection status.",
      messages: [{ role: "user", content: "hello" }],
      tools: [
        {
          name: "range.read",
          description: "Read an Excel range",
          parameters: {
            type: "object",
            properties: {
              sheetName: { type: "string" },
              range: { type: "string" },
            },
          },
          riskLevel: "safe",
          requiresApproval: false,
        },
      ],
    });

    expect(withPromptAndTools).toBeGreaterThan(messageOnly);
    expect(withPromptAndTools).toBeGreaterThan(estimateTokens("hello"));
  });
});

// ============================================================
// shouldCompact
// ============================================================

describe("shouldCompact", () => {
  it("should return false when disabled", () => {
    const config: CompactionConfig = { ...DEFAULT_COMPACTION_CONFIG, enabled: false };
    expect(shouldCompact([], config)).toBe(false);
  });

  it("should return false for empty items", () => {
    expect(shouldCompact([], DEFAULT_COMPACTION_CONFIG)).toBe(false);
  });

  it("should return true when tokens exceed threshold", () => {
    const config: CompactionConfig = {
      ...DEFAULT_COMPACTION_CONFIG,
      autoCompactTokenThreshold: 10, // very low threshold
    };
    const items: TurnItem[] = [
      {
        type: "user_message",
        id: "1",
        content: "这是一条足够长的消息来触发压缩阈值测试",
        timestamp: Date.now(),
      },
      {
        type: "assistant_message",
        id: "2",
        content: "这是助手的回复内容，也需要足够的长度来触发阈值",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    expect(shouldCompact(items, config)).toBe(true);
  });
});

// ============================================================
// collectUserMessages
// ============================================================

describe("collectUserMessages", () => {
  it("should collect only user messages", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
      {
        type: "assistant_message",
        id: "2",
        content: "world",
        phase: "final",
        timestamp: Date.now(),
      },
      { type: "user_message", id: "3", content: "again", timestamp: Date.now() },
    ];
    expect(collectUserMessages(items)).toHaveLength(2);
  });

  it("should exclude summary messages", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
      {
        type: "user_message",
        id: "2",
        content: `${SUMMARY_PREFIX}\n摘要内容`,
        timestamp: Date.now(),
      },
    ];
    expect(collectUserMessages(items)).toHaveLength(1);
  });
});

// ============================================================
// buildCompactedHistory
// ============================================================

describe("buildCompactedHistory", () => {
  it("should include summary message", () => {
    const userMessages = [
      { type: "user_message" as const, id: "1", content: "hello", timestamp: Date.now() },
    ];
    const result = buildCompactedHistory(userMessages, "测试摘要");
    const summaryMsg = result.find(
      (item) => item.type === "user_message" && item.content.startsWith(SUMMARY_PREFIX),
    );
    expect(summaryMsg).toBeDefined();
  });

  it("should retain recent messages within budget", () => {
    const userMessages = [
      { type: "user_message" as const, id: "1", content: "short", timestamp: Date.now() },
    ];
    const result = buildCompactedHistory(userMessages, "摘要");
    // 短消息应在预算内
    const nonSummary = result.filter(
      (item) => item.type === "user_message" && !item.content.startsWith(SUMMARY_PREFIX),
    );
    expect(nonSummary.length).toBeGreaterThanOrEqual(1);
  });

  it("should keep only configured recent user messages", () => {
    const userMessages = [
      { type: "user_message" as const, id: "1", content: "第一条", timestamp: 1 },
      { type: "user_message" as const, id: "2", content: "第二条", timestamp: 2 },
      { type: "user_message" as const, id: "3", content: "第三条", timestamp: 3 },
    ];
    const result = buildCompactedHistory(userMessages, "摘要", {
      ...DEFAULT_COMPACTION_CONFIG,
      retainedRecentItemCount: 2,
      retainedUserMessageMaxTokens: 10_000,
    });
    const retained = result.filter(
      (item) => item.type === "user_message" && !item.content.startsWith(SUMMARY_PREFIX),
    );

    expect(retained.map((item) => item.id)).toEqual(["2", "3"]);
  });
});

// ============================================================
// performCompaction
// ============================================================

describe("performCompaction", () => {
  it("should return compactedItem and newHistory", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
      {
        type: "assistant_message",
        id: "2",
        content: "world",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    const result = performCompaction(items, "测试摘要", "auto_pre_turn");
    expect(result.compactedItem).toBeDefined();
    expect(result.compactedItem.type).toBe("compacted");
    expect(result.compactedItem.summary).toBe("测试摘要");
    expect(result.compactedItem.reason).toBe("auto_pre_turn");
    expect(result.newHistory).toBeDefined();
    expect(Array.isArray(result.newHistory)).toBe(true);
  });

  it("should record tokensBefore and tokensAfter", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
    ];
    const result = performCompaction(items, "摘要", "auto_token_limit");
    expect(result.compactedItem.tokensBefore).toBeGreaterThan(0);
    expect(result.compactedItem.tokensAfter).toBeGreaterThan(0);
  });
});

// ============================================================
// historyToCompactPrompt
// ============================================================

describe("historyToCompactPrompt", () => {
  it("should format user messages", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "帮我处理数据", timestamp: Date.now() },
    ];
    const prompt = historyToCompactPrompt(items);
    expect(prompt).toContain("【用户】");
    expect(prompt).toContain("帮我处理数据");
  });

  it("should format assistant messages", () => {
    const items: TurnItem[] = [
      {
        type: "assistant_message",
        id: "1",
        content: "好的",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    const prompt = historyToCompactPrompt(items);
    expect(prompt).toContain("【助手】");
  });

  it("should format tool calls", () => {
    const items: TurnItem[] = [
      {
        type: "tool_call",
        id: "1",
        toolName: "range.read",
        arguments: { range: "A1" },
        status: "completed",
        timestamp: Date.now(),
      },
    ];
    const prompt = historyToCompactPrompt(items);
    expect(prompt).toContain("【工具调用】");
  });

  it("should format tool results", () => {
    const items: TurnItem[] = [
      {
        type: "tool_result",
        id: "1",
        toolCallId: "tc-1",
        toolName: "range.read",
        result: "output",
        isError: false,
        timestamp: Date.now(),
      },
    ];
    const prompt = historyToCompactPrompt(items);
    expect(prompt).toContain("【工具结果】");
  });
});

// ============================================================
// buildResumeContext
// ============================================================

describe("buildResumeContext", () => {
  it("should include [中断恢复上下文] header", () => {
    const items: TurnItem[] = [
      { type: "user_message", id: "1", content: "hello", timestamp: Date.now() },
      {
        type: "assistant_message",
        id: "2",
        content: "world",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    const ctx = buildResumeContext(items);
    expect(ctx).toContain("[中断恢复上下文]");
  });

  it("should include last assistant message", () => {
    const items: TurnItem[] = [
      {
        type: "assistant_message",
        id: "1",
        content: "first",
        phase: "final",
        timestamp: Date.now(),
      },
      {
        type: "assistant_message",
        id: "2",
        content: "last reply",
        phase: "final",
        timestamp: Date.now(),
      },
    ];
    const ctx = buildResumeContext(items);
    expect(ctx).toContain("last reply");
  });

  it("should handle empty items", () => {
    const ctx = buildResumeContext([]);
    expect(ctx).toContain("[中断恢复上下文]");
    expect(ctx).toContain("请基于以上上下文继续之前的工作");
  });
});
