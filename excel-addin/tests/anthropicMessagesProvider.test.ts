import { describe, expect, it, vi } from "vitest";
import { AnthropicMessagesStreamProvider } from "../shared/provider/anthropicMessagesProvider";
import { encodeAnthropicMessagesBody } from "../shared/provider/anthropicMessagesEncode";
import { buildToolNameMaps, isToolNameMaps } from "../shared/provider/openaiToolNameMap";
import type { AgentStreamEvent, StreamChatRequest } from "../shared/agent/types";
import type { ToolDefinition } from "../shared/tools/types";

const tools: ToolDefinition[] = [
  {
    name: "host.status",
    description: "status",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, additionalProperties: false },
  },
  {
    name: "range.read",
    description: "read",
    riskLevel: "safe",
    parameters: {
      type: "object",
      properties: {
        sheetName: { type: "string" },
        range: { type: "string" },
      },
      additionalProperties: false,
    },
  },
];

function req(overrides?: Partial<StreamChatRequest>): StreamChatRequest {
  return {
    systemPrompt: "You are helpful.",
    messages: [{ role: "user", content: "hi" }],
    tools,
    ...overrides,
  };
}

function sseResponse(chunks: string[]): Response {
  const enc = new TextEncoder();
  let i = 0;
  return new Response(
    new ReadableStream<Uint8Array>({
      pull(c) {
        if (i >= chunks.length) {
          c.close();
          return;
        }
        c.enqueue(enc.encode(chunks[i++]));
      },
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } },
  );
}

function data(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

async function collect(
  provider: AnthropicMessagesStreamProvider,
  request: StreamChatRequest = req(),
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(request)) out.push(e);
  return out;
}

function provider(fetchImpl: ReturnType<typeof vi.fn>, maxTokens?: number) {
  return new AnthropicMessagesStreamProvider({
    baseUrl: "https://api.anthropic.com/v1",
    apiKey: "sk-ant-secret",
    model: "claude-3-5-sonnet-latest",
    fetchImpl,
    ...(maxTokens != null ? { maxTokens } : {}),
  });
}

describe("Anthropic Messages encode", () => {
  it("skips history system; merges consecutive tool_result; maps tools", () => {
    const maps = buildToolNameMaps(tools);
    expect(isToolNameMaps(maps)).toBe(true);
    if (!isToolNameMaps(maps)) return;
    const encoded = encodeAnthropicMessagesBody(
      "sys-top",
      [
        { role: "system", content: "ignored-history" },
        { role: "user", content: "q" },
        {
          role: "assistant",
          content: "calling",
          toolCalls: [
            { id: "tu1", name: "host.status", argumentsJson: "{}" },
            { id: "tu2", name: "range.read", argumentsJson: '{"sheetName":"S"}' },
          ],
        },
        { role: "tool", toolCallId: "tu1", name: "host.status", content: '{"ok":true}' },
        { role: "tool", toolCallId: "tu2", name: "range.read", content: '{"v":1}' },
        { role: "user", content: "next" },
      ],
      tools,
      maps,
    );
    expect("error" in encoded).toBe(false);
    if ("error" in encoded) return;
    expect(encoded.system).toBe("sys-top");
    expect(encoded.messages).toEqual([
      { role: "user", content: "q" },
      {
        role: "assistant",
        content: [
          { type: "text", text: "calling" },
          { type: "tool_use", id: "tu1", name: "host_status", input: {} },
          {
            type: "tool_use",
            id: "tu2",
            name: "range_read",
            input: { sheetName: "S" },
          },
        ],
      },
      {
        role: "user",
        content: [
          { type: "tool_result", tool_use_id: "tu1", content: '{"ok":true}' },
          { type: "tool_result", tool_use_id: "tu2", content: '{"v":1}' },
        ],
      },
      { role: "user", content: "next" },
    ]);
    expect(encoded.tools[0]).toEqual({
      name: "host_status",
      description: "status",
      input_schema: tools[0]?.parameters,
    });
  });

  it("rejects bad tool input / empty id before fetch", async () => {
    const maps = buildToolNameMaps(tools);
    expect(isToolNameMaps(maps)).toBe(true);
    if (!isToolNameMaps(maps)) return;
    const badJson = encodeAnthropicMessagesBody(
      "",
      [
        { role: "user", content: "q" },
        {
          role: "assistant",
          content: "",
          toolCalls: [{ id: "x", name: "host.status", argumentsJson: "not-json" }],
        },
      ],
      tools,
      maps,
    );
    expect(badJson).toMatchObject({ error: expect.stringContaining("not valid JSON") });

    const fetchImpl = vi.fn();
    const events = await collect(
      provider(fetchImpl),
      req({
        messages: [
          { role: "user", content: "q" },
          {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "  ", name: "host.status", argumentsJson: "{}" }],
          },
        ],
      }),
    );
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
  });
});

describe("AnthropicMessagesStreamProvider happy paths", () => {
  it("headers x-api-key only; body max_tokens/system/tools; text + stop", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        data({
          type: "message_start",
          message: { usage: { input_tokens: 5, output_tokens: 0 } },
        }),
        data({
          type: "content_block_start",
          index: 0,
          content_block: { type: "text", text: "" },
        }),
        data({
          type: "content_block_delta",
          index: 0,
          delta: { type: "text_delta", text: "Hi" },
        }),
        data({ type: "content_block_stop", index: 0 }),
        data({
          type: "message_delta",
          delta: { stop_reason: "end_turn" },
          usage: { output_tokens: 2 },
        }),
        data({ type: "message_stop" }),
        "data: [DONE]\n\n",
      ]),
    );
    const events = await collect(provider(fetchImpl));
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe("https://api.anthropic.com/v1/messages");
    const headers = init.headers as Record<string, string>;
    expect(headers["x-api-key"]).toBe("sk-ant-secret");
    expect(headers["anthropic-version"]).toBe("2023-06-01");
    expect(headers.Accept).toBe("text/event-stream");
    expect(headers.Authorization).toBeUndefined();
    expect(JSON.stringify(headers)).not.toContain("anthropic-dangerous");
    const body = JSON.parse(String(init.body));
    expect(body).toMatchObject({
      model: "claude-3-5-sonnet-latest",
      max_tokens: 4096,
      stream: true,
      system: "You are helpful.",
    });
    expect(body.tools[0].name).toBe("host_status");
    expect(body.tools[0].input_schema).toEqual(tools[0]?.parameters);
    expect(JSON.stringify(body)).not.toContain("sk-ant-secret");
    expect(events.filter((e) => e.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Hi" },
    ]);
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "stop",
    });
  });

  it("two tools by index; multi text; invalid JSON args still end", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        data({
          type: "message_start",
          message: { usage: { input_tokens: 1, output_tokens: 0 } },
        }),
        data({
          type: "content_block_start",
          index: 0,
          content_block: { type: "text", text: "A" },
        }),
        data({
          type: "content_block_delta",
          index: 0,
          delta: { type: "text_delta", text: "B" },
        }),
        data({ type: "content_block_stop", index: 0 }),
        data({
          type: "content_block_start",
          index: 1,
          content_block: { type: "tool_use", id: "call_a", name: "host_status" },
        }),
        data({
          type: "content_block_start",
          index: 2,
          content_block: { type: "tool_use", id: "call_b", name: "range_read" },
        }),
        data({
          type: "content_block_delta",
          index: 2,
          delta: { type: "input_json_delta", partial_json: '{"r":' },
        }),
        data({
          type: "content_block_delta",
          index: 1,
          delta: { type: "input_json_delta", partial_json: "not-json" },
        }),
        data({
          type: "content_block_delta",
          index: 2,
          delta: { type: "input_json_delta", partial_json: "1}" },
        }),
        data({ type: "content_block_stop", index: 1 }),
        data({ type: "content_block_stop", index: 2 }),
        data({
          type: "message_delta",
          delta: { stop_reason: "tool_use" },
          usage: { output_tokens: 4 },
        }),
        data({ type: "message_stop" }),
        "data: [DONE]\n\n",
      ]),
    );
    const events = await collect(provider(fetchImpl));
    expect(
      events
        .filter((e) => e.type === "text_delta")
        .map((e) => (e as { delta: string }).delta),
    ).toEqual(["A", "B"]);
    expect(events.filter((e) => e.type === "tool_call_begin")).toEqual([
      { type: "tool_call_begin", toolCallId: "call_a", toolName: "host.status" },
      { type: "tool_call_begin", toolCallId: "call_b", toolName: "range.read" },
    ]);
    expect(events.filter((e) => e.type === "tool_call_end")).toEqual([
      {
        type: "tool_call_end",
        toolCallId: "call_a",
        toolName: "host.status",
        argumentsJson: "not-json",
      },
      {
        type: "tool_call_end",
        toolCallId: "call_b",
        toolName: "range.read",
        argumentsJson: '{"r":1}',
      },
    ]);
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "tool_calls",
    });
  });

  it("usage start input + delta output → last snapshot has both", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        data({
          type: "message_start",
          message: { usage: { input_tokens: 10, output_tokens: 0 } },
        }),
        data({
          type: "content_block_start",
          index: 0,
          content_block: { type: "text", text: "x" },
        }),
        data({ type: "content_block_stop", index: 0 }),
        data({
          type: "message_delta",
          delta: { stop_reason: "end_turn" },
          usage: { output_tokens: 7 },
        }),
        data({ type: "message_stop" }),
        "data: [DONE]\n\n",
      ]),
    );
    const events = await collect(provider(fetchImpl));
    const usages = events.filter((e) => e.type === "usage") as Array<{
      type: "usage";
      usage: { inputTokens: number; outputTokens: number };
    }>;
    expect(usages.length).toBeGreaterThanOrEqual(1);
    const last = usages.at(-1)!;
    expect(last.usage.inputTokens).toBe(10);
    expect(last.usage.outputTokens).toBe(7);
  });
});
