import { describe, expect, it } from "vitest";
import { collectAgentStream } from "../shared/agent/collectStream";
import type { AgentStreamEvent } from "../shared/agent/types";

async function* of(...events: AgentStreamEvent[]) {
  for (const e of events) yield e;
}

describe("collectAgentStream", () => {
  it("aggregates text + finish stop", async () => {
    const r = await collectAgentStream(
      of(
        { type: "text_delta", delta: "Hello" },
        { type: "text_delta", delta: " world" },
        { type: "finish", reason: "stop" },
      ),
    );
    expect(r.assistantText).toBe("Hello world");
    expect(r.toolCalls).toEqual([]);
    expect(r.finishReason).toBe("stop");
    expect(r.error).toBeUndefined();
  });

  it("interleaves multi-id calls and keeps first-seen order", async () => {
    const r = await collectAgentStream(
      of(
        { type: "tool_call_begin", toolCallId: "b", toolName: "range.read" },
        { type: "tool_call_begin", toolCallId: "a", toolName: "host.status" },
        { type: "tool_call_delta", toolCallId: "a", argumentsDelta: "{}" },
        { type: "tool_call_delta", toolCallId: "b", argumentsDelta: '{"x":1}' },
        { type: "tool_call_end", toolCallId: "a" },
        { type: "tool_call_end", toolCallId: "b" },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(r.toolCalls.map((c) => c.id)).toEqual(["b", "a"]);
    expect(r.toolCalls[0]?.name).toBe("range.read");
    expect(r.toolCalls[1]?.name).toBe("host.status");
  });

  it("end non-empty argumentsJson wins over delta", async () => {
    const r = await collectAgentStream(
      of(
        { type: "tool_call_begin", toolCallId: "1", toolName: "t" },
        { type: "tool_call_delta", toolCallId: "1", argumentsDelta: '{"from":"delta"}' },
        {
          type: "tool_call_end",
          toolCallId: "1",
          argumentsJson: '{"from":"end"}',
        },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(r.toolCalls[0]?.argumentsJson).toBe('{"from":"end"}');
  });

  it("supports end-only and delta-before-begin", async () => {
    const endOnly = await collectAgentStream(
      of(
        {
          type: "tool_call_end",
          toolCallId: "e",
          toolName: "host.status",
          argumentsJson: "{}",
        },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(endOnly.toolCalls).toEqual([
      { id: "e", name: "host.status", argumentsJson: "{}" },
    ]);

    const deltaFirst = await collectAgentStream(
      of(
        { type: "tool_call_delta", toolCallId: "d", argumentsDelta: '{"a":1}' },
        { type: "tool_call_begin", toolCallId: "d", toolName: "range.read" },
        { type: "tool_call_end", toolCallId: "d" },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(deltaFirst.toolCalls[0]).toEqual({
      id: "d",
      name: "range.read",
      argumentsJson: '{"a":1}',
    });
  });

  it("ignores duplicate end", async () => {
    const r = await collectAgentStream(
      of(
        { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
        { type: "tool_call_end", toolCallId: "1", argumentsJson: '{"a":1}' },
        { type: "tool_call_end", toolCallId: "1", argumentsJson: '{"a":2}' },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(r.toolCalls).toHaveLength(1);
    expect(r.toolCalls[0]?.argumentsJson).toBe('{"a":1}');
  });

  it("discards missing end", async () => {
    const r = await collectAgentStream(
      of(
        { type: "tool_call_begin", toolCallId: "1", toolName: "host.status" },
        { type: "tool_call_delta", toolCallId: "1", argumentsDelta: "{}" },
        { type: "finish", reason: "tool_calls" },
      ),
    );
    expect(r.toolCalls).toEqual([]);
  });

  it("usage last event is authoritative snapshot", async () => {
    const r = await collectAgentStream(
      of(
        { type: "usage", usage: { inputTokens: 1, outputTokens: 2 } },
        { type: "text_delta", delta: "x" },
        {
          type: "usage",
          usage: { inputTokens: 10, outputTokens: 20, cachedInputTokens: 3 },
        },
        { type: "finish", reason: "stop" },
      ),
    );
    expect(r.usage).toEqual({
      inputTokens: 10,
      outputTokens: 20,
      cachedInputTokens: 3,
    });
  });

  it("error returns without requiring finish; abort throws", async () => {
    const errored = await collectAgentStream(
      of(
        { type: "text_delta", delta: "partial" },
        { type: "error", message: "boom", kind: "http", status: 500 },
        { type: "finish", reason: "stop" },
      ),
    );
    expect(errored.error).toEqual({
      message: "boom",
      kind: "http",
      status: 500,
      url: undefined,
    });
    // Still returns partial text aggregation up to error break
    expect(errored.assistantText).toBe("partial");

    const controller = new AbortController();
    controller.abort();
    await expect(collectAgentStream(of({ type: "finish", reason: "stop" }), controller.signal)).rejects.toMatchObject({
      name: "AbortError",
    });
  });

  it("defaults finish to stop when stream ends cleanly without finish", async () => {
    const r = await collectAgentStream(of({ type: "text_delta", delta: "ok" }));
    expect(r.finishReason).toBe("stop");
  });

  it("empty end argumentsJson falls back to delta then {}", async () => {
    const withDelta = await collectAgentStream(
      of(
        { type: "tool_call_delta", toolCallId: "1", argumentsDelta: '{"d":1}' },
        { type: "tool_call_end", toolCallId: "1", toolName: "t", argumentsJson: "" },
      ),
    );
    expect(withDelta.toolCalls[0]?.argumentsJson).toBe('{"d":1}');

    const empty = await collectAgentStream(
      of({ type: "tool_call_end", toolCallId: "2", toolName: "t" }),
    );
    expect(empty.toolCalls[0]?.argumentsJson).toBe("{}");
  });
});
