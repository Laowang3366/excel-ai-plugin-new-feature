import { describe, expect, it, vi } from "vitest";
import { OpenAIChatCompletionsStreamProvider } from "../shared/provider/openaiChatCompletionsProvider";
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

function sseResponse(chunks: string[], init?: ResponseInit): Response {
  const encoder = new TextEncoder();
  let i = 0;
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (i >= chunks.length) {
        controller.close();
        return;
      }
      controller.enqueue(encoder.encode(chunks[i++]));
    },
  });
  return new Response(stream, {
    status: 200,
    headers: { "Content-Type": "text/event-stream" },
    ...init,
  });
}

async function collect(
  provider: OpenAIChatCompletionsStreamProvider,
  req: StreamChatRequest,
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(req)) out.push(e);
  return out;
}

function baseReq(overrides: Partial<StreamChatRequest> = {}): StreamChatRequest {
  return {
    systemPrompt: "You are helpful.",
    messages: [{ role: "user", content: "hi" }],
    tools,
    ...overrides,
  };
}

describe("OpenAIChatCompletionsStreamProvider", () => {
  it("streams multi text deltas over SSE chunk/CRLF boundaries then stop", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        'data: {"choices":[{"delta":{"content":"Hel"}}]}\r\n\r\n',
        'data: {"choices":[{"delta":{"content":"lo"}}]}\n\n',
        'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
        "data: [DONE]\n\n",
      ]),
    );
    const provider = new OpenAIChatCompletionsStreamProvider({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-test",
      model: "gpt-4o",
      fetchImpl,
    });
    const events = await collect(provider, baseReq());
    expect(events.filter((e) => e.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "Hel" },
      { type: "text_delta", delta: "lo" },
    ]);
    expect(events.some((e) => e.type === "finish" && e.reason === "stop")).toBe(true);
    expect(events.some((e) => e.type === "error")).toBe(false);
    const firstCall = fetchImpl.mock.calls[0] as unknown as [string, RequestInit] | undefined;
    const body = JSON.parse(String(firstCall?.[1]?.body));
    expect(body.stream).toBe(true);
    expect(body.model).toBe("gpt-4o");
    expect(body.messages[0]).toEqual({ role: "system", content: "You are helpful." });
    expect(body.tools.map((t: { function: { name: string } }) => t.function.name)).toEqual([
      "host_status",
      "range_read",
    ]);
    expect(JSON.stringify(body)).not.toContain("sk-test");
  });

  it("interleaves two tool_calls with id/name/arguments fragments", async () => {
    const fetchImpl = vi.fn(async () =>
      sseResponse([
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_a","function":{"name":"host_status"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":1,"id":"call_b","function":{"name":"range_read"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":"{\\"x\\""}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":1,"function":{"arguments":"{\\"sheetName\\":\\"S\\""}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"arguments":":1}"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{"tool_calls":[{"index":1,"function":{"arguments":",\\"range\\":\\"A1\\"}"}}]}}]}\n\n',
        'data: {"choices":[{"delta":{},"finish_reason":"tool_calls"}]}\n\n',
        "data: [DONE]\n\n",
      ]),
    );
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl,
      }),
      baseReq(),
    );
    const nonError = events.filter((e) => e.type !== "error");
    // Begin only when first arguments fragment arrives (id+name already complete).
    expect(nonError.map((e) => e.type)).toEqual([
      "tool_call_begin",
      "tool_call_delta",
      "tool_call_begin",
      "tool_call_delta",
      "tool_call_delta",
      "tool_call_delta",
      "tool_call_end",
      "tool_call_end",
      "finish",
    ]);
    expect(nonError[0]).toEqual({
      type: "tool_call_begin",
      toolCallId: "call_a",
      toolName: "host.status",
    });
    expect(nonError[2]).toEqual({
      type: "tool_call_begin",
      toolCallId: "call_b",
      toolName: "range.read",
    });
    const ends = nonError.filter((e) => e.type === "tool_call_end");
    expect(ends[0]).toMatchObject({
      toolCallId: "call_a",
      toolName: "host.status",
      argumentsJson: '{"x":1}',
    });
    expect(ends[1]).toMatchObject({
      toolCallId: "call_b",
      toolName: "range.read",
      argumentsJson: '{"sheetName":"S","range":"A1"}',
    });
    expect(nonError.at(-1)).toEqual({ type: "finish", reason: "tool_calls" });
  });

  it("maps function names and encodes assistant/tool history", async () => {
    const maps = buildToolNameMaps(tools);
    expect(isToolNameMaps(maps)).toBe(true);
    if (!isToolNameMaps(maps)) return;
    expect(maps.internalToExternal.get("host.status")).toBe("host_status");
    expect(
      isToolNameMaps(
        buildToolNameMaps([
          { name: "a.b" as ToolDefinition["name"], description: "", riskLevel: "safe", parameters: {} },
          { name: "a_b" as ToolDefinition["name"], description: "", riskLevel: "safe", parameters: {} },
        ] as ToolDefinition[]),
      ),
    ).toBe(false);

    let captured: Record<string, unknown> | undefined;
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      captured = JSON.parse(String(init?.body));
      return sseResponse([
        'data: {"choices":[{"delta":{},"finish_reason":"stop"}]}\n\n',
        "data: [DONE]\n\n",
      ]);
    });
    await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://example.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl,
      }),
      baseReq({
        messages: [
          { role: "user", content: "do" },
          {
            role: "assistant",
            content: "",
            toolCalls: [{ id: "c1", name: "host.status", argumentsJson: "{}" }],
          },
          {
            role: "tool",
            content: '{"ok":true}',
            toolCallId: "c1",
            name: "host.status",
          },
        ],
      }),
    );
    expect(captured?.messages).toEqual([
      { role: "system", content: "You are helpful." },
      { role: "user", content: "do" },
      {
        role: "assistant",
        content: null,
        tool_calls: [
          {
            id: "c1",
            type: "function",
            function: { name: "host_status", arguments: "{}" },
          },
        ],
      },
      { role: "tool", tool_call_id: "c1", content: '{"ok":true}' },
    ]);
  });

  it("parse errors: malformed JSON, missing id/name, EOF without finish", async () => {
    const malformed = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => sseResponse(["data: {not-json}\n\n"])),
      }),
      baseReq(),
    );
    expect(malformed.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(malformed.some((e) => e.type === "finish")).toBe(false);

    const missingId = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"function":{"name":"host_status","arguments":"{}"}}]},"finish_reason":"tool_calls"}]}\n\n',
            "data: [DONE]\n\n",
          ]),
        ),
      }),
      baseReq(),
    );
    expect(missingId.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const noFinish = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse(['data: {"choices":[{"delta":{"content":"x"}}]}\n\n', "data: [DONE]\n\n"]),
        ),
      }),
      baseReq(),
    );
    expect(noFinish.at(-1)).toMatchObject({ type: "error", kind: "parse" });
  });

  it("HTTP/CORS/network/abort/empty key/empty body/usage", async () => {
    const noKey = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "",
        model: "m",
        fetchImpl: vi.fn(),
      }),
      baseReq(),
    );
    expect(noKey[0]).toMatchObject({ type: "error", kind: "missing_key" });

    const http = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk-secret",
        model: "m",
        fetchImpl: vi.fn(
          async () =>
            new Response(JSON.stringify({ error: { message: "bad key" } }), { status: 401 }),
        ),
      }),
      baseReq(),
    );
    expect(http[0]).toMatchObject({ type: "error", kind: "http", status: 401 });
    expect(JSON.stringify(http)).not.toContain("sk-secret");

    const cors = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => {
          throw new TypeError("Failed to fetch");
        }),
      }),
      baseReq(),
    );
    expect(cors[0]).toMatchObject({ type: "error", kind: "cors" });

    const network = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => {
          throw new Error("ECONNRESET");
        }),
      }),
      baseReq(),
    );
    expect(network[0]).toMatchObject({ type: "error", kind: "network" });

    const pre = new AbortController();
    pre.abort();
    const abortedPre = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(),
      }),
      baseReq({ signal: pre.signal }),
    );
    expect(abortedPre[0]).toMatchObject({ type: "error", kind: "aborted" });
    const abortedFetch = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => {
          throw Object.assign(new Error("aborted"), { name: "AbortError" });
        }),
      }),
      baseReq(),
    );
    expect(abortedFetch[0]).toMatchObject({ type: "error", kind: "aborted" });

    const emptyBody = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(
          async () =>
            new Response(null, { status: 200, headers: { "Content-Type": "text/event-stream" } }),
        ),
      }),
      baseReq(),
    );
    expect(emptyBody[0]).toMatchObject({ type: "error", kind: "parse" });

    const usageEvents = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            'data: {"choices":[{"delta":{"content":"ok"},"finish_reason":"stop"}]}\n\n',
            'data: {"choices":[],"usage":{"prompt_tokens":11,"completion_tokens":2,"prompt_tokens_details":{"cached_tokens":1},"completion_tokens_details":{"reasoning_tokens":0}}}\n\n',
            "data: [DONE]\n\n",
          ]),
        ),
      }),
      baseReq(),
    );
    const types = usageEvents.map((e) => e.type);
    expect(types).toContain("usage");
    expect(types.indexOf("usage")).toBeLessThan(types.lastIndexOf("finish"));
    expect(types.filter((x) => x !== "error").at(-1)).toBe("finish");
    const usage = usageEvents.find((e) => e.type === "usage");
    expect(usage).toMatchObject({
      type: "usage",
      usage: { inputTokens: 11, outputTokens: 2, cachedInputTokens: 1 },
    });
  });

  it("tool message without toolCallId fails before fetch", async () => {
    const fetchImpl = vi.fn();
    const events = await collect(
      new OpenAIChatCompletionsStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl,
      }),
      baseReq({
        messages: [{ role: "tool", content: "{}", name: "host.status" }],
      }),
    );
    expect(events[0]).toMatchObject({ type: "error", kind: "parse" });
    expect(fetchImpl).not.toHaveBeenCalled();
  });

});
