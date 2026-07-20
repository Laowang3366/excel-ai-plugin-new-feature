import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  estimateRequestTokens,
  resolveOutputReserve,
} from "../shared/agent/historyBudget";
import {
  ScriptedStreamProvider,
  textThenStop,
  toolCallThenFinish,
} from "../shared/agent/scriptedProvider";
import type { AgentMessage, StreamChatRequest } from "../shared/agent/types";
import type { ToolDefinition, ToolName, ToolResult } from "../shared/tools/types";

function fakeExecutor() {
  return {
    execute: vi.fn(
      async (call: {
        name: ToolName;
        arguments: Record<string, unknown>;
      }): Promise<ToolResult> => ({
        ok: true,
        tool: call.name,
        data: { name: call.name },
      }),
    ),
  };
}

/** English ≈ 4 chars/token — size to force trimming under small windows. */
function bulk(label: string, tokens: number): string {
  return `${label}|${"w".repeat(tokens * 4)}`;
}

function largeUser(label: string, tokens = 4_000): AgentMessage {
  return { role: "user", content: bulk(label, tokens) };
}

function largeAssistant(label: string, tokens = 4_000): AgentMessage {
  return { role: "assistant", content: bulk(label, tokens) };
}

const statusTool: ToolDefinition = {
  name: "host.status",
  description: "s",
  riskLevel: "safe",
  parameters: { type: "object", properties: {} },
};

function hardLimit(ctx: number): number {
  return ctx - resolveOutputReserve(ctx);
}

function assertNoOrphanTools(msgs: AgentMessage[]): void {
  for (let i = 0; i < msgs.length; i += 1) {
    const m = msgs[i];
    if (m.role !== "tool") continue;
    let j = i - 1;
    while (j >= 0 && msgs[j].role === "tool") j -= 1;
    expect(msgs[j]?.role).toBe("assistant");
    expect(msgs[j]?.toolCalls?.some((c) => c.id === m.toolCallId)).toBe(true);
  }
}

function assertNoOrphanAssistant(msgs: AgentMessage[]): void {
  for (let i = 0; i < msgs.length; i += 1) {
    if (msgs[i].role !== "assistant") continue;
    let j = i - 1;
    while (j >= 0 && msgs[j].role !== "user") j -= 1;
    expect(j).toBeGreaterThanOrEqual(0);
    expect(msgs[j].role).toBe("user");
  }
}

function assertRequestHistoryWellFormed(msgs: AgentMessage[]): void {
  assertNoOrphanTools(msgs);
  assertNoOrphanAssistant(msgs);
}

describe("AgentLoop request history budget", () => {
  it("does not trim when contextWindowSize is unset", async () => {
    const history: AgentMessage[] = [largeUser("h0", 500), largeAssistant("h1", 500)];
    const historySnapshot = structuredClone(history);
    const provider = new ScriptedStreamProvider({
      rounds: [textThenStop("ok")],
    });
    const result = await new AgentLoop({
      provider,
      executor: fakeExecutor(),
      systemPrompt: "s",
      tools: [],
    }).run({ userMessage: "now", history });

    expect(provider.lastRequest?.messages.map((m) => m.content.split("|")[0])).toEqual([
      "h0",
      "h1",
      "now",
    ]);
    expect(result.messages).toHaveLength(4);
    expect(history).toEqual(historySnapshot);
  });

  it("trims only streamChat request copy; keeps full result history; does not mutate input", async () => {
    const history: AgentMessage[] = [
      largeUser("OLD0"),
      largeAssistant("OLD1"),
      largeUser("OLD2"),
      largeAssistant("OLD3"),
    ];
    const historySnapshot = structuredClone(history);
    const requests: StreamChatRequest[] = [];
    const ctx = 12_000;
    const provider = new ScriptedStreamProvider({
      rounds: [
        (streamCtx) => {
          requests.push({
            systemPrompt: streamCtx.request.systemPrompt,
            messages: streamCtx.request.messages.slice(),
            tools: streamCtx.request.tools.slice(),
          });
          return textThenStop("final");
        },
      ],
    });

    const result = await new AgentLoop({
      provider,
      executor: fakeExecutor(),
      systemPrompt: "sys",
      tools: [],
      contextWindowSize: ctx,
    }).run({ userMessage: "current-turn", history });

    expect(requests).toHaveLength(1);
    const reqMsgs = requests[0].messages;
    expect(reqMsgs.some((m) => m.content === "current-turn")).toBe(true);
    expect(reqMsgs.some((m) => m.content.startsWith("OLD0|"))).toBe(false);
    expect(
      estimateRequestTokens({
        systemPrompt: "sys",
        tools: [],
        messages: reqMsgs,
      }),
    ).toBeLessThanOrEqual(hardLimit(ctx));

    expect(result.messages.some((m) => m.content.startsWith("OLD0|"))).toBe(true);
    expect(result.messages.some((m) => m.content === "current-turn")).toBe(true);
    expect(result.messages[result.messages.length - 1]).toMatchObject({
      role: "assistant",
      content: "final",
    });
    expect(history).toEqual(historySnapshot);
  });

  it("applies budget on every agent round request copy", async () => {
    const history: AgentMessage[] = [
      largeUser("OLD0"),
      largeAssistant("OLD1"),
      largeUser("OLD2"),
      largeAssistant("OLD3"),
    ];
    const requests: AgentMessage[][] = [];
    const ctx = 14_000;
    const provider = new ScriptedStreamProvider({
      rounds: [
        (streamCtx) => {
          requests.push(streamCtx.request.messages.slice());
          return toolCallThenFinish("c1", "host.status", "{}");
        },
        (streamCtx) => {
          requests.push(streamCtx.request.messages.slice());
          return textThenStop("done");
        },
      ],
    });

    const result = await new AgentLoop({
      provider,
      executor: fakeExecutor(),
      systemPrompt: "sys",
      tools: [statusTool],
      contextWindowSize: ctx,
      maxRounds: 4,
    }).run({ userMessage: "ask-tools", history });

    expect(requests).toHaveLength(2);
    for (const msgs of requests) {
      expect(msgs.some((m) => m.content === "ask-tools")).toBe(true);
      assertRequestHistoryWellFormed(msgs);
      expect(
        estimateRequestTokens({
          systemPrompt: "sys",
          tools: [statusTool],
          messages: msgs,
        }),
      ).toBeLessThanOrEqual(hardLimit(ctx));
    }

    const round2 = requests[1];
    expect(round2.some((m) => m.role === "assistant" && m.toolCalls?.length)).toBe(
      true,
    );
    expect(round2.some((m) => m.role === "tool" && m.toolCallId === "c1")).toBe(true);

    expect(result.messages.some((m) => m.content.startsWith("OLD0|"))).toBe(true);
    expect(result.messages.filter((m) => m.role === "tool")).toHaveLength(1);
    expect(result.status).toBe("completed");
  });

  it("keeps tool chain atomic when dropping old history with tools", async () => {
    const history: AgentMessage[] = [
      { role: "user", content: bulk("h-u", 5_000) },
      {
        role: "assistant",
        content: "",
        toolCalls: [{ id: "old", name: "host.status", argumentsJson: "{}" }],
      },
      {
        role: "tool",
        content: JSON.stringify({ ok: true, data: bulk("tool", 5_000) }),
        toolCallId: "old",
        name: "host.status",
      },
      { role: "assistant", content: bulk("after", 5_000) },
    ];
    const provider = new ScriptedStreamProvider({
      rounds: [textThenStop("ok")],
    });
    await new AgentLoop({
      provider,
      executor: fakeExecutor(),
      systemPrompt: "s",
      tools: [],
      contextWindowSize: 10_000,
    }).run({ userMessage: "now", history });

    const msgs = provider.lastRequest!.messages;
    assertRequestHistoryWellFormed(msgs);
    const hasOldTool = msgs.some((m) => m.toolCallId === "old");
    const hasOldAssistant = msgs.some(
      (m) => m.role === "assistant" && m.toolCalls?.some((c) => c.id === "old"),
    );
    expect(hasOldTool).toBe(hasOldAssistant);
    expect(msgs.some((m) => m.content === "now")).toBe(true);
  });
});
