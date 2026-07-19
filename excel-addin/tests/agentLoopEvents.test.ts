import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import { collectAgentStream } from "../shared/agent/collectStream";
import {
  ScriptedStreamProvider,
  errorEvent,
  textThenStop,
} from "../shared/agent/scriptedProvider";
import { throwIfAborted } from "../shared/agent/streamProvider";
import type { AgentStreamEvent, LoopEvent } from "../shared/agent/types";
import { ToolExecutor } from "../shared/tools";
import type { ToolName, ToolResult } from "../shared/tools/types";
import { MockHostAdapter } from "./mockHost";

async function* of(...events: AgentStreamEvent[]) {
  for (const e of events) yield e;
}

describe("agent loop event hardening", () => {
  it("collect onTextDelta fires per delta even if stream later errors", async () => {
    const seen: string[] = [];
    const r = await collectAgentStream(
      of(
        { type: "text_delta", delta: "A" },
        { type: "text_delta", delta: "B" },
        { type: "error", message: "boom", kind: "provider" },
      ),
      { onTextDelta: (d) => seen.push(d) },
    );
    expect(seen).toEqual(["A", "B"]);
    expect(r.assistantText).toBe("AB");
    expect(r.error?.message).toBe("boom");
  });

  it("loop emits two text_delta events before round_end; assistantText joins", async () => {
    const events: LoopEvent[] = [];
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "text_delta", delta: "Hello" },
          { type: "text_delta", delta: "!" },
          { type: "finish", reason: "stop" },
        ],
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "s",
      onEvent: (e) => events.push(e),
    }).run({ userMessage: "u" });

    const deltas = events.filter((e) => e.type === "text_delta");
    expect(deltas).toEqual([
      { type: "text_delta", delta: "Hello", round: 1 },
      { type: "text_delta", delta: "!", round: 1 },
    ]);
    const endIdx = events.findIndex((e) => e.type === "round_end");
    const lastDeltaIdx = events.map((e) => e.type).lastIndexOf("text_delta");
    expect(lastDeltaIdx).toBeGreaterThan(-1);
    expect(lastDeltaIdx).toBeLessThan(endIdx);
    expect(result.assistantText).toBe("Hello!");
    expect(result.status).toBe("completed");
  });

  it("tool_outcome carries toolCallId for each tool", async () => {
    const outcomes: { id: string; name: string }[] = [];
    const executor = {
      execute: vi.fn(async (call: { name: ToolName; arguments: Record<string, unknown> }) => {
        return { ok: true, tool: call.name, data: call.name } satisfies ToolResult;
      }),
    };
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "tool_call_begin", toolCallId: "id-a", toolName: "host.status" },
          { type: "tool_call_end", toolCallId: "id-a", argumentsJson: "{}" },
          { type: "tool_call_begin", toolCallId: "id-b", toolName: "sheet.list" },
          { type: "tool_call_end", toolCallId: "id-b", argumentsJson: "{}" },
          { type: "finish", reason: "tool_calls" },
        ],
        textThenStop("done"),
      ],
    });
    await new AgentLoop({
      provider,
      executor,
      systemPrompt: "s",
      onEvent: (e) => {
        if (e.type === "tool_outcome") {
          outcomes.push({ id: e.toolCallId, name: e.outcome.toolName });
        }
      },
    }).run({ userMessage: "u" });
    expect(outcomes).toEqual([
      { id: "id-a", name: "host.status" },
      { id: "id-b", name: "sheet.list" },
    ]);
  });

  it("custom abort reason becomes AbortError / aborted status", async () => {
    const c = new AbortController();
    c.abort(new Error("stop-now"));
    expect(() => throwIfAborted(c.signal)).toThrow(
      expect.objectContaining({ name: "AbortError", message: "stop-now" }),
    );

    const provider = new ScriptedStreamProvider({
      rounds: [[{ type: "text_delta", delta: "x" }, { type: "finish", reason: "stop" }]],
      eventDelayMs: 25,
    });
    const mid = new AbortController();
    const orig = provider.streamChat.bind(provider);
    provider.streamChat = async function* (req) {
      let n = 0;
      for await (const e of orig(req)) {
        yield e;
        n += 1;
        if (n === 1) mid.abort(new Error("custom-mid"));
      }
    };
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "s",
      signal: mid.signal,
    }).run({ userMessage: "u" });
    expect(result.status).toBe("aborted");
    expect(result.lastFinishReason).toBe("aborted");
    expect(result.rounds).toBe(0);
  });

  it("provider failure sets lastFinishReason error; tools slice to provider", async () => {
    const tools = [
      {
        name: "host.status" as const,
        description: "d",
        riskLevel: "safe" as const,
        parameters: { type: "object", properties: {}, additionalProperties: false },
      },
    ];
    const provider = new ScriptedStreamProvider({
      rounds: [errorEvent("down", "http", 500)],
    });
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "s",
      tools,
    }).run({ userMessage: "u" });
    expect(result.status).toBe("failed");
    expect(result.lastFinishReason).toBe("error");
    expect(provider.lastRequest?.tools).not.toBe(tools);
    expect(provider.lastRequest?.tools).toEqual(tools);
  });

  it("mid-stream deltas still observe before failed error round", async () => {
    const deltas: string[] = [];
    const provider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "text_delta", delta: "p1" },
          { type: "text_delta", delta: "p2" },
          { type: "error", message: "die", kind: "provider" },
        ],
      ],
    });
    const result = await new AgentLoop({
      provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "s",
      onEvent: (e) => {
        if (e.type === "text_delta") deltas.push(e.delta);
      },
    }).run({ userMessage: "u" });
    expect(deltas).toEqual(["p1", "p2"]);
    expect(result.status).toBe("failed");
    expect(result.rounds).toBe(0);
    expect(result.assistantText).toBe("");
    expect(result.messages.some((m) => m.role === "assistant")).toBe(false);
    expect(result.lastFinishReason).toBe("error");
  });
});
