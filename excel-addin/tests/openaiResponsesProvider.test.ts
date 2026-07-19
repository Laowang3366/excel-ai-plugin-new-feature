import { describe, expect, it, vi } from "vitest";
import { OpenAIResponsesStreamProvider } from "../shared/provider/openaiResponsesProvider";
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
      properties: { sheetName: { type: "string" }, range: { type: "string" } },
      additionalProperties: false,
    },
  },
];

function req(overrides: Partial<StreamChatRequest> = {}): StreamChatRequest {
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

async function collect(
  provider: OpenAIResponsesStreamProvider,
  request: StreamChatRequest = req(),
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(request)) out.push(e);
  return out;
}

function data(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

describe("OpenAIResponsesStreamProvider", () => {
  it("streams text delta + done suffix + completed stop", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        data({
          type: "response.output_text.delta",
          item_id: "msg_1",
          content_index: 0,
          delta: "Hel",
        }),
        data({
          type: "response.output_text.delta",
          item_id: "msg_1",
          content_index: 0,
          delta: "lo",
        }),
        data({
          type: "response.output_text.done",
          item_id: "msg_1",
          content_index: 0,
          text: "Hello!",
        }),
        data({
          type: "response.completed",
          response: {
            usage: {
              input_tokens: 3,
              output_tokens: 2,
              input_tokens_details: { cached_tokens: 1 },
            },
          },
        }),
        "data: [DONE]\n\n",
      ]),
    );
    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk-test",
        model: "gpt-4o",
        fetchImpl,
      }),
    );
    const deltas = events.filter((e) => e.type === "text_delta");
    expect(deltas.map((e) => (e as { delta: string }).delta)).toEqual(["Hel", "lo", "!"]);
    expect(events.some((e) => e.type === "usage")).toBe(true);
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "stop",
    });
    const firstCall = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(firstCall[0]).toBe("https://api.openai.com/v1/responses");
    const body = JSON.parse(String(firstCall[1].body));
    expect(body).toMatchObject({ model: "gpt-4o", stream: true });
    expect(body.instructions).toBe("You are helpful.");
    expect(body.tools[0]).toEqual({
      type: "function",
      name: "host_status",
      description: "status",
      parameters: tools[0]?.parameters,
    });
    expect(JSON.stringify(body)).not.toContain("sk-test");
  });

  it("single tool: added → args deltas → args.done override → item.done → completed", async () => {
    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({
              type: "response.output_item.added",
              item: {
                type: "function_call",
                id: "fc_item",
                call_id: "call_abc",
                name: "host_status",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "fc_item",
              delta: '{"x":',
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "fc_item",
              delta: "1}",
            }),
            data({
              type: "response.function_call_arguments.done",
              item_id: "fc_item",
              arguments: '{"x":9}',
            }),
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "fc_item",
                call_id: "call_abc",
                name: "host_status",
                arguments: '{"final":true}',
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    const begin = events.find((e) => e.type === "tool_call_begin");
    const end = events.find((e) => e.type === "tool_call_end");
    expect(begin).toEqual({
      type: "tool_call_begin",
      toolCallId: "call_abc",
      toolName: "host.status",
    });
    expect(end).toMatchObject({
      toolCallId: "call_abc",
      toolName: "host.status",
      argumentsJson: '{"final":true}',
    });
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "tool_calls",
    });
  });

  it("two item_ids interleave without cross-talk", async () => {
    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({
              type: "response.output_item.added",
              item: {
                type: "function_call",
                id: "i1",
                call_id: "c1",
                name: "host_status",
              },
            }),
            data({
              type: "response.output_item.added",
              item: {
                type: "function_call",
                id: "i2",
                call_id: "c2",
                name: "range_read",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i2",
              delta: '{"r":1}',
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i1",
              delta: "{}",
            }),
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "i1",
                call_id: "c1",
                name: "host_status",
                arguments: "{}",
              },
            }),
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "i2",
                call_id: "c2",
                name: "range_read",
                arguments: '{"r":1}',
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    const ends = events.filter((e) => e.type === "tool_call_end");
    expect(ends).toEqual([
      expect.objectContaining({
        toolCallId: "c1",
        toolName: "host.status",
        argumentsJson: "{}",
      }),
      expect.objectContaining({
        toolCallId: "c2",
        toolName: "range.read",
        argumentsJson: '{"r":1}',
      }),
    ]);
  });

  it("item.done without prior added still begin/end", async () => {
    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "only",
                call_id: "call_x",
                name: "host_status",
                arguments: "{}",
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.find((e) => e.type === "tool_call_begin")).toMatchObject({
      toolCallId: "call_x",
      toolName: "host.status",
    });
    expect(events.find((e) => e.type === "tool_call_end")).toMatchObject({
      toolCallId: "call_x",
      argumentsJson: "{}",
    });
  });

  it("encodes second-round function_call + function_call_output with same call_id", async () => {
    let captured: Record<string, unknown> | undefined;
    await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async (_u, init) => {
          captured = JSON.parse(String(init?.body));
          return sseResponse([
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]);
        }),
      }),
      req({
        messages: [
          { role: "user", content: "do" },
          {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "call_1", name: "host.status", argumentsJson: "{}" }],
          },
          {
            role: "tool",
            toolCallId: "call_1",
            content: '{"ok":true}',
            name: "host.status",
          },
        ],
      }),
    );
    expect(captured?.input).toEqual([
      { type: "message", role: "user", content: "do" },
      {
        type: "function_call",
        call_id: "call_1",
        name: "host_status",
        arguments: "{}",
        status: "completed",
      },
      {
        type: "function_call_output",
        call_id: "call_1",
        output: '{"ok":true}',
      },
    ]);
  });

  it("maps host.status and detects collision", () => {
    const maps = buildToolNameMaps(tools);
    expect(isToolNameMaps(maps)).toBe(true);
    if (isToolNameMaps(maps)) {
      expect(maps.internalToExternal.get("host.status")).toBe("host_status");
    }
  });
});
