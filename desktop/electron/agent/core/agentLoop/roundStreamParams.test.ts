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

  it("passes the current user intent into advanced Office tool visibility", async () => {
    const execute = async () => ({ success: true });
    const toolExecutors = new Map([
      ["office.action.apply", { name: "office.action.apply", execute }],
      ["office.action.inspect", { name: "office.action.inspect", execute }],
      ["office.workflow.run", { name: "office.workflow.run", execute }],
    ]);
    const simple = await buildRoundStreamParams({
      turnItemGroups: [[userItem("把 A1 写入 1")]],
      turnInput: { content: "把 A1 写入 1" },
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
      },
      toolExecutors,
      round: 1,
    });
    const pivot = await buildRoundStreamParams({
      turnItemGroups: [[userItem("创建数据透视表")]],
      turnInput: { content: "创建数据透视表" },
      aiConfig: {
        provider: "openai",
        apiKey: "test",
        baseUrl: "https://example.com/v1",
        model: "gpt-test",
      },
      toolExecutors,
      round: 1,
    });

    expect(getToolOperations(simple.toolDefs, "office.action.apply"))
      .not.toContain("createPivotTable");
    expect(getToolOperations(pivot.toolDefs, "office.action.apply"))
      .toContain("createPivotTable");
  });
});

function getToolOperations(
  definitions: Awaited<ReturnType<typeof buildRoundStreamParams>>["toolDefs"],
  toolName: string,
): string[] {
  const definition = definitions.find((tool) => tool.name === toolName);
  const properties = definition?.parameters.properties as Record<string, any> | undefined;
  return properties?.operation?.enum ?? [];
}
