import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  ScriptedStreamProvider,
  textThenStop,
} from "../shared/agent/scriptedProvider";
import type { ToolName, ToolResult } from "../shared/tools/types";
import { ToolExecutor } from "../shared/tools";
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

describe("AgentLoop abort paths", () => {
  it("abort before sample / mid-stream / before tool / during tool", async () => {
    // before sample
    const c1 = new AbortController();
    c1.abort();
    let r = await new AgentLoop({
      provider: new ScriptedStreamProvider({ rounds: [textThenStop("x")] }),
      executor: hostExecutor(),
      systemPrompt: "s",
      signal: c1.signal,
    }).run({ userMessage: "u" });
    expect(r.status).toBe("aborted");
    expect(r.rounds).toBe(0);
    expect(r.messages.some((m) => m.role === "assistant")).toBe(false);

    // mid-stream
    const c2 = new AbortController();
    const midProvider = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "text_delta", delta: "partial" },
          { type: "finish", reason: "stop" },
        ],
      ],
      eventDelayMs: 20,
    });
    // Abort after first event via wrapping: abort during delay of second event
    const orig = midProvider.streamChat.bind(midProvider);
    midProvider.streamChat = async function* (req) {
      let i = 0;
      for await (const e of orig(req)) {
        yield e;
        i += 1;
        if (i === 1) c2.abort();
      }
    };
    r = await new AgentLoop({
      provider: midProvider,
      executor: hostExecutor(),
      systemPrompt: "s",
      signal: c2.signal,
    }).run({ userMessage: "u" });
    expect(r.status).toBe("aborted");
    expect(r.rounds).toBe(0);

    // before second tool
    const c3 = new AbortController();
    const order: string[] = [];
    const exec3 = fakeExecutor(async (name) => {
      order.push(name);
      if (name === "host.status") c3.abort();
      return { ok: true, tool: name as ToolName, data: true };
    });
    r = await new AgentLoop({
      provider: new ScriptedStreamProvider({
        rounds: [
          [
            { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
            { type: "tool_call_end", toolCallId: "1", argumentsJson: "{}" },
            { type: "tool_call_begin", toolCallId: "2", toolName: "sheet.list" },
            { type: "tool_call_end", toolCallId: "2", argumentsJson: "{}" },
            { type: "finish", reason: "tool_calls" },
          ],
          textThenStop("nope"),
        ],
      }),
      executor: exec3,
      systemPrompt: "s",
      signal: c3.signal,
    }).run({ userMessage: "u" });
    expect(r.status).toBe("aborted");
    expect(order).toEqual(["host.status"]);
    expect(r.messages.filter((m) => m.role === "tool")).toHaveLength(1);

    // during tool: settle, write message, stop subsequent
    const c4 = new AbortController();
    let resolveExec: (() => void) | undefined;
    const gate = new Promise<void>((res) => {
      resolveExec = res;
    });
    const exec4 = fakeExecutor(async (name) => {
      if (name === "host.status") {
        c4.abort();
        await gate;
      }
      return { ok: true, tool: name as ToolName, data: name };
    });
    const runP = new AgentLoop({
      provider: new ScriptedStreamProvider({
        rounds: [
          [
            { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
            { type: "tool_call_end", toolCallId: "1", argumentsJson: "{}" },
            { type: "tool_call_begin", toolCallId: "2", toolName: "sheet.list" },
            { type: "tool_call_end", toolCallId: "2", argumentsJson: "{}" },
            { type: "finish", reason: "tool_calls" },
          ],
          textThenStop("no"),
        ],
      }),
      executor: exec4,
      systemPrompt: "s",
      signal: c4.signal,
    }).run({ userMessage: "u" });
    await new Promise((r) => setTimeout(r, 5));
    resolveExec?.();
    r = await runP;
    expect(r.status).toBe("aborted");
    expect(r.messages.filter((m) => m.role === "tool")).toHaveLength(1);
    expect(exec4.execute).toHaveBeenCalledTimes(1);
  });

});
