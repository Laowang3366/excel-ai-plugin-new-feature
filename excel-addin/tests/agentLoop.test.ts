import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  ScriptedStreamProvider,
  errorEvent,
  textThenStop,
  toolCallThenFinish,
} from "../shared/agent/scriptedProvider";
import type { AgentStreamEvent, LoopEvent } from "../shared/agent/types";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import type { ToolName, ToolResult } from "../shared/tools/types";
import { MockHostAdapter } from "./mockHost";

function hostExecutor() {
  return new ToolExecutor(new MockHostAdapter());
}

function fakeExecutor(impl: (name: string, args: Record<string, unknown>) => Promise<ToolResult> | ToolResult) {
  return {
    execute: vi.fn(async (call: { name: ToolName; arguments: Record<string, unknown> }) => {
      return await impl(call.name, call.arguments);
    }),
  };
}

function eventsForTool(
  id: string,
  name: string,
  argsJson: string,
  finish: "tool_calls" | "stop" | "length" = "tool_calls",
): AgentStreamEvent[] {
  return [
    { type: "tool_call_begin", toolCallId: id, toolName: name },
    { type: "tool_call_delta", toolCallId: id, argumentsDelta: argsJson },
    { type: "tool_call_end", toolCallId: id, toolName: name, argumentsJson: argsJson },
    { type: "finish", reason: finish },
  ];
}

describe("AgentLoop", () => {
  it("completes one text-only round without tools", async () => {
    const provider = new ScriptedStreamProvider({ rounds: [textThenStop("done")] });
    const loop = new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "sys",
    });
    const result = await loop.run({ userMessage: "hi" });
    expect(result.status).toBe("completed");
    expect(result.rounds).toBe(1);
    expect(result.assistantText).toBe("done");
    expect(result.messages.map((m) => m.role)).toEqual(["user", "assistant"]);
    expect(result.lastFinishReason).toBe("stop");
  });

  it("tool success then final answer; second request has assistant+tool", async () => {
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish(
          "c1",
          "host.status",
          "{}",
        ),
        textThenStop("final"),
      ],
    });
    const loop = new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "S",
    });
    const result = await loop.run({ userMessage: "go" });
    expect(result.status).toBe("completed");
    expect(result.rounds).toBe(2);
    expect(result.assistantText).toBe("final");
    expect(result.messages.map((m) => m.role)).toEqual([
      "user",
      "assistant",
      "tool",
      "assistant",
    ]);
    const second = provider.lastRequest;
    expect(second?.messages.map((m) => m.role)).toEqual(["user", "assistant", "tool"]);
    expect(second?.systemPrompt).toBe("S");
    const toolMsg = result.messages.find((m) => m.role === "tool");
    expect(toolMsg?.toolCallId).toBe("c1");
    expect(toolMsg?.content).toContain('"kind":"host"');
  });

  it("executes two tools in one round strictly serially", async () => {
    const order: string[] = [];
    const executor = fakeExecutor(async (name) => {
      order.push(`start:${name}`);
      await new Promise((r) => setTimeout(r, 5));
      order.push(`end:${name}`);
      return { ok: true, tool: name as ToolName, data: {} };
    });
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "1", toolName: "host.status", argumentsJson: "{}" },
          { type: "tool_call_begin", toolCallId: "2", toolName: "sheet.list" },
          { type: "tool_call_end", toolCallId: "2", toolName: "sheet.list", argumentsJson: "{}" },
          { type: "finish", reason: "tool_calls" },
        ],
        textThenStop("ok"),
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor,
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(result.status).toBe("completed");
    expect(order).toEqual([
      "start:host.status",
      "end:host.status",
      "start:sheet.list",
      "end:sheet.list",
    ]);
    expect(result.messages.filter((m) => m.role === "tool")).toHaveLength(2);
  });

  it("unsupported and ordinary failure feed back; run completes", async () => {
    const executor = fakeExecutor(async (name) => {
      if (name === "host.status") {
        return {
          ok: false,
          tool: "host.status",
          unsupported: true,
          error: "nope",
        };
      }
      return { ok: false, tool: name as ToolName, error: "ordinary fail" };
    });
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "1", argumentsJson: "{}" },
          { type: "tool_call_begin", toolCallId: "2", toolName: "sheet.list" },
          { type: "tool_call_end", toolCallId: "2", argumentsJson: "{}" },
          { type: "finish", reason: "tool_calls" },
        ],
        textThenStop("after-fail"),
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor,
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(result.status).toBe("completed");
    const tools = result.messages.filter((m) => m.role === "tool");
    expect(tools[0]?.content).toContain('"unsupported":true');
    expect(tools[1]?.content).toContain("ordinary fail");
  });

  it("unknown/bad JSON/array/null args never call executor", async () => {
    const executor = fakeExecutor(async () => ({
      ok: true,
      tool: "host.status",
      data: {},
    }));
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "u", toolName: "not.a.tool" },
          { type: "tool_call_end", toolCallId: "u", argumentsJson: "{}" },
          { type: "tool_call_begin", toolCallId: "b", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "b", argumentsJson: "{bad" },
          { type: "tool_call_begin", toolCallId: "a", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "a", argumentsJson: "[]" },
          { type: "tool_call_begin", toolCallId: "n", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "n", argumentsJson: "null" },
          { type: "finish", reason: "tool_calls" },
        ],
        textThenStop("done"),
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor,
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(executor.execute).not.toHaveBeenCalled();
    expect(result.status).toBe("completed");
    const assistant = result.messages.find((m) => m.role === "assistant" && m.toolCalls);
    expect(assistant?.toolCalls?.map((c) => c.id)).toEqual(["u", "b", "a", "n"]);
    const kinds = result.messages
      .filter((m) => m.role === "tool")
      .map((m) => JSON.parse(m.content).kind);
    expect(kinds).toEqual([
      "unknown_tool",
      "invalid_arguments",
      "invalid_arguments",
      "invalid_arguments",
    ]);
  });

  it("maxRounds=1 runs tools then max_rounds without resample", async () => {
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish("1", "host.status", "{}"),
        textThenStop("should-not-run"),
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "s",
      maxRounds: 1,
    }).run({ userMessage: "u" });
    expect(result.status).toBe("max_rounds");
    expect(result.rounds).toBe(1);
    expect(provider.callCount).toBe(1);
    expect(result.messages.filter((m) => m.role === "tool")).toHaveLength(1);
    expect(result.lastFinishReason).toBe("max_rounds");
  });

  it("provider error/throw => failed without partial assistant", async () => {
    const errProvider = new ScriptedStreamProvider({
      rounds: [errorEvent("down", "http", 503)],
    });
    let r = await new AgentLoop({
      provider: errProvider,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(r.status).toBe("failed");
    expect(r.rounds).toBe(0);
    expect(r.messages.some((m) => m.role === "assistant")).toBe(false);
    expect(r.error?.message).toBe("down");
    expect(r.lastFinishReason).toBe("error");

    const throwProvider = {
      async *streamChat() {
        throw new Error("network boom");
      },
    };
    r = await new AgentLoop({
      provider: throwProvider,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(r.status).toBe("failed");
    expect(r.rounds).toBe(0);
    expect(r.error?.message).toBe("network boom");
    expect(r.lastFinishReason).toBe("error");
  });

  it("finish does not gate tools; length without tools completes", async () => {
    const p1 = new ScriptedStreamProvider({
      rounds: [
        eventsForTool("1", "host.status", "{}", "stop"),
        textThenStop("after"),
      ],
    });
    let r = await new AgentLoop({
      provider: p1,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(r.status).toBe("completed");
    expect(r.messages.some((m) => m.role === "tool")).toBe(true);

    const p2 = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "x", toolName: "host.status" },
          { type: "finish", reason: "tool_calls" },
        ],
      ],
    });
    r = await new AgentLoop({
      provider: p2,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(r.status).toBe("completed");
    expect(r.rounds).toBe(1);
    expect(r.lastFinishReason).toBe("tool_calls");

    const p3 = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "text_delta", delta: "cut" },
          { type: "finish", reason: "length" },
        ],
      ],
    });
    r = await new AgentLoop({
      provider: p3,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(r.status).toBe("completed");
    expect(r.lastFinishReason).toBe("length");
  });

  it("sums usage across rounds including optional fields", async () => {
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "usage", usage: { inputTokens: 1, outputTokens: 2, cachedInputTokens: 1 } },
          ...toolCallThenFinish("1", "host.status", "{}"),
        ],
        [
          {
            type: "usage",
            usage: {
              inputTokens: 3,
              outputTokens: 4,
              cachedInputTokens: 2,
              reasoningOutputTokens: 5,
            },
          },
          { type: "text_delta", delta: "end" },
          { type: "finish", reason: "stop" },
        ],
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(result.usage).toEqual({
      inputTokens: 4,
      outputTokens: 6,
      cachedInputTokens: 3,
      reasoningOutputTokens: 5,
    });
  });

  it("systemPrompt/tools/history; custom subset blocks inactive tools", async () => {
    const subset = TOOL_DEFINITIONS.filter((t) => t.name === "host.status");
    const executor = fakeExecutor(async (name) => ({
      ok: true,
      tool: name as ToolName,
      data: name,
    }));
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "1", toolName: "sheet.list" },
          { type: "tool_call_end", toolCallId: "1", argumentsJson: "{}" },
          { type: "finish", reason: "tool_calls" },
        ],
        textThenStop("ok"),
      ],
    });
    const history = [{ role: "user" as const, content: "prev" }];
    const result = await new AgentLoop({
      provider,
      executor,
      systemPrompt: "PROMPT",
      tools: subset,
    }).run({ userMessage: "now", history });
    expect(provider.lastRequest?.systemPrompt).toBe("PROMPT");
    expect(result.messages[0]).toEqual({ role: "user", content: "prev" });
    expect(executor.execute).not.toHaveBeenCalled();
    expect(result.messages.find((m) => m.role === "tool")?.content).toContain("unknown_tool");
  });

  it("onEvent order and single run_end", async () => {
    const events: LoopEvent["type"][] = [];
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish("1", "host.status", "{}"),
        textThenStop("bye"),
      ],
    });
    await new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "s",
      onEvent: (e) => events.push(e.type),
    }).run({ userMessage: "u" });
    expect(events[0]).toBe("round_start");
    expect(events.filter((t) => t === "run_end")).toEqual(["run_end"]);
    expect(events.at(-1)).toBe("run_end");
    expect(events).toContain("tool_call_parsed");
    expect(events).toContain("tool_outcome");
    expect(events).toContain("round_end");
  });

  it("no fetch during agent loop", async () => {
    const spy = vi.spyOn(globalThis, "fetch");
    const provider = new ScriptedStreamProvider({ rounds: [textThenStop("x")] });
    await new AgentLoop({
      provider,
      executor: hostExecutor(),
      systemPrompt: "s",
    }).run({ userMessage: "u" });
    expect(spy).not.toHaveBeenCalled();
    spy.mockRestore();
  });
});
