import { describe, expect, it } from "vitest";

import type { TurnItem } from "../../shared/types";
import { buildRoundStreamParams } from "./roundStreamParams";

const userItem = (content: string): TurnItem => ({
  type: "user_message",
  id: "u1",
  content,
  timestamp: 1,
});

describe("buildRoundStreamParams", () => {
  it("builds stream params with resume context and configured reasoning mode", async () => {
    const result = await buildRoundStreamParams({
      turnItemGroups: [[userItem("继续分析")]],
      turnInput: { content: "继续分析" },
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
      },
      configuredReasoningMode: "medium",
      baseSystemPrompt: "系统提示",
      round: 3,
      resumeContext: "从中断处恢复",
    });

    expect(result.streamParams.messages).toEqual([
      expect.objectContaining({ role: "user", content: "继续分析" }),
      { role: "system", content: "从中断处恢复" },
    ]);
    expect(result.streamParams.reasoningMode).toBe("medium");
    expect(result.streamParams.roundId).toBe(3);
    expect(result.streamParams.tools).toEqual([]);
    expect(result.effectiveSystemPrompt).toContain("系统提示");
  });

  it("lets aiConfig reasoning mode override configured fallback", async () => {
    const result = await buildRoundStreamParams({
      turnItemGroups: [[userItem("测试")]],
      turnInput: { content: "测试" },
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
        reasoningMode: "low",
      },
      configuredReasoningMode: "high",
      round: 1,
    });

    expect(result.streamParams.reasoningMode).toBe("low");
  });

  it("keeps formula workflow guidance when the user answers a clarification", async () => {
    const previous: TurnItem[] = [
      userItem("【功能模块：生成公式】\n数据源选区：Sheet1!A1:B10\n答案填入锚点/选区：Sheet1!D1"),
      { type: "tool_call", id: "prepare", toolName: "formula.prepare", arguments: {}, status: "completed", timestamp: 2 },
      {
        type: "tool_result",
        id: "result-prepare",
        toolCallId: "prepare",
        toolName: "formula.prepare",
        result: {
          status: "needs_clarification",
          scenario: "分组聚合",
          inputShape: "记录表",
          outputShape: "汇总表",
          inputGrain: "明细",
          outputGrain: "部门",
          businessKeys: ["部门"],
          transformChain: [],
          constraints: [],
          acceptanceChecks: [],
          assumptions: [],
          clarificationQuestion: "重复项是否合计？",
        },
        isError: false,
        timestamp: 3,
      },
    ];
    const current = [userItem("重复项需要合计")];

    const result = await buildRoundStreamParams({
      turnItemGroups: [previous, current],
      turnInput: { content: "重复项需要合计" },
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
      },
      round: 1,
    });

    expect(result.effectiveSystemPrompt).toContain("场景化操作指南：公式助手");
    expect(result.effectiveSystemPrompt).toContain("formula.prepare");
  });
});
