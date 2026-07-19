import { describe, expect, it, vi } from "vitest";
import { OpenAIChatCompletionsStreamProvider } from "../shared/provider/openaiChatCompletionsProvider";
import { SseByteParser } from "../shared/provider/openaiSse";
import type { AgentStreamEvent, StreamChatRequest } from "../shared/agent/types";
import type { ToolDefinition } from "../shared/tools/types";

const tools: ToolDefinition[] = [
  {
    name: "host.status",
    description: "status",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, additionalProperties: false },
  },
];

function req(overrides: Partial<StreamChatRequest> = {}): StreamChatRequest {
  return {
    systemPrompt: "sys",
    messages: [{ role: "user", content: "hi" }],
    tools,
    ...overrides,
  };
}

function responseFromBytes(chunks: Uint8Array[]): Response {
  let i = 0;
  return new Response(
    new ReadableStream<Uint8Array>({
      pull(controller) {
        if (i >= chunks.length) {
          controller.close();
          return;
        }
        controller.enqueue(chunks[i++]);
      },
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } },
  );
}

async function collect(
  provider: OpenAIChatCompletionsStreamProvider,
  request: StreamChatRequest = req(),
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(request)) out.push(e);
  return out;
}

describe("OpenAI SSE / tool-call boundaries", () => {
  it("keeps multi-byte UTF-8 text across arbitrary byte splits", async () => {
    const text = "你好世界";
    const payload = `data: {"choices":[{"delta":{"content":"${text}"}}]}\n\ndata: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\ndata: [DONE]\n\n`;
    const bytes = new TextEncoder().encode(payload);
    // Split inside the first multi-byte character of 你 (UTF-8 is 3 bytes).
    const cut = payload.indexOf(text) + 1; // after first byte of 你
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => responseFromBytes([bytes.slice(0, cut), bytes.slice(cut)])),
      }),
    );
    const deltas = events.filter((e) => e.type === "text_delta");
    expect(deltas.map((e) => (e as { delta: string }).delta).join("")).toBe(text);
    expect(events.some((e) => e.type === "error")).toBe(false);
  });

  it("does not treat split CRLF as blank line; multi data lines join", () => {
    const parser = new SseByteParser();
    // JSON split across two data: lines with CRLF; event terminator CRLF split across chunks.
    const part1 = new TextEncoder().encode('data: {"choices":[{"delta":{"content":"A"');
    const part2 = new TextEncoder().encode('}}]}\r');
    const part3 = new TextEncoder().encode('\n\r\n');
    expect(parser.push(part1)).toEqual([]);
    expect(parser.push(part2)).toEqual([]); // lone CR held
    const mid = parser.push(part3);
    expect(mid).toEqual([{ kind: "data", data: '{"choices":[{"delta":{"content":"A"}}]}' }]);
  });

  it("flush emits finish event at EOF without [DONE] or trailing blank line", async () => {
    const payload =
      'data: {"choices":[{"delta":{"content":"ok"},"finish_reason":"stop"}]}';
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          responseFromBytes([new TextEncoder().encode(payload)]),
        ),
      }),
    );
    expect(events.some((e) => e.type === "text_delta" && e.delta === "ok")).toBe(true);
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "stop",
    });
  });

  it("appends id/name fragments and delays begin until complete", async () => {
    const chunks = [
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_"}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"1","function":{"name":"host_"}}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"status","arguments":"{\\"a\\":"}}]}}]}\n\n',
      'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"1}"}}]},"finish_reason":"tool_calls"}]}\n\n',
      "data: [DONE]\n\n",
    ];
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          responseFromBytes(chunks.map((c) => new TextEncoder().encode(c))),
        ),
      }),
    );
    const types = events.map((e) => e.type);
    const beginAt = types.indexOf("tool_call_begin");
    const firstDelta = types.indexOf("tool_call_delta");
    expect(beginAt).toBeGreaterThan(-1);
    expect(firstDelta).toBeGreaterThan(beginAt);
    expect(events[beginAt]).toEqual({
      type: "tool_call_begin",
      toolCallId: "call_1",
      toolName: "host.status",
    });
    const end = events.find((e) => e.type === "tool_call_end");
    expect(end).toMatchObject({
      toolCallId: "call_1",
      toolName: "host.status",
      argumentsJson: '{"a":1}',
    });
  });

  it("index-only tool slot without id/name is parse error at finish", async () => {
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          responseFromBytes([
            new TextEncoder().encode(
              'data: {"choices":[{"delta":{"tool_calls":[{"index":0}]},"finish_reason":"tool_calls"}]}\n\ndata: [DONE]\n\n',
            ),
          ]),
        ),
      }),
    );
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(events.some((e) => e.type === "finish")).toBe(false);
  });

  it("unknown final external name is parse error", async () => {
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          responseFromBytes([
            new TextEncoder().encode(
              'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"c1","function":{"name":"not_a_tool","arguments":"{}"}}]},"finish_reason":"tool_calls"}]}\n\ndata: [DONE]\n\n',
            ),
          ]),
        ),
      }),
    );
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
  });

  it("blank historical tool call id fails before fetch", async () => {
    const fetchImpl = vi.fn();
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl,
      }),
      req({
        messages: [
          {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "   ", name: "host.status", argumentsJson: "{}" }],
          },
        ],
      }),
    );
    expect(events[0]).toMatchObject({ type: "error", kind: "parse" });
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("redacts apiKey from HTTP error bodies", async () => {
    const key = "sk-super-secret-key";
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: key,
        model: "m",
        fetchImpl: vi.fn(
          async () =>
            new Response(JSON.stringify({ error: { message: `invalid ${key}` } }), {
              status: 401,
            }),
        ),
      }),
    );
    expect(events[0]).toMatchObject({ type: "error", kind: "http", status: 401 });
    expect(JSON.stringify(events)).not.toContain(key);
    expect(JSON.stringify(events)).toContain("[REDACTED]");
  });
});
