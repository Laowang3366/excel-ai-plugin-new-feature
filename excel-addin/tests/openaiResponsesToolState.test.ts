import { describe, expect, it, vi } from "vitest";
import { OpenAIResponsesStreamProvider } from "../shared/provider/openaiResponsesProvider";
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

function req(): StreamChatRequest {
  return {
    systemPrompt: "sys",
    messages: [{ role: "user", content: "hi" }],
    tools,
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
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(req())) out.push(e);
  return out;
}

function data(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}

`;
}

describe("OpenAI Responses tool-state hardening", () => {
  it("call_id conflict after freeze is parse error without wrong end", async () => {
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
                call_id: "call_a",
                name: "host_status",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i1",
              delta: "{}",
            }),
            data({
              type: "response.function_call_arguments.done",
              item_id: "i1",
              call_id: "call_b",
              arguments: "{}",
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(events.some((e) => e.type === "finish")).toBe(false);
    expect(events.some((e) => e.type === "tool_call_end")).toBe(false);
  });

  it("name conflict is parse error", async () => {
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
                call_id: "call_a",
                name: "host_status",
              },
            }),
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "i1",
                call_id: "call_a",
                name: "range_read",
                arguments: "{}",
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    // range_read not in tools map either, but conflict should fire first if name differs
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(events.some((e) => e.type === "finish")).toBe(false);
  });

  it("added+delta without any done cannot end at completed", async () => {
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
                call_id: "call_a",
                name: "host_status",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i1",
              delta: '{"x":1}',
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(events.some((e) => e.type === "finish")).toBe(false);
    expect(events.some((e) => e.type === "tool_call_end")).toBe(false);
  });

  it("args.done without item.done may complete at terminal", async () => {
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
                call_id: "call_a",
                name: "host_status",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i1",
              delta: '{"x":',
            }),
            data({
              type: "response.function_call_arguments.done",
              item_id: "i1",
              call_id: "call_a",
              name: "host_status",
              arguments: '{"x":2}',
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.find((e) => e.type === "tool_call_end")).toMatchObject({
      toolCallId: "call_a",
      toolName: "host.status",
      argumentsJson: '{"x":2}',
    });
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "tool_calls",
    });
  });

  it("item.done empty arguments string wins over prior deltas", async () => {
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
                call_id: "call_a",
                name: "host_status",
              },
            }),
            data({
              type: "response.function_call_arguments.delta",
              item_id: "i1",
              delta: '{"junk":true}',
            }),
            data({
              type: "response.output_item.done",
              item: {
                type: "function_call",
                id: "i1",
                call_id: "call_a",
                name: "host_status",
                arguments: "",
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.find((e) => e.type === "tool_call_end")).toMatchObject({
      toolCallId: "call_a",
      toolName: "host.status",
      argumentsJson: "{}",
    });
  });

  it("official response.failed nests error under response and redacts key", async () => {
    const key = "sk-super-secret";
    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: key,
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({
              type: "response.failed",
              response: {
                status: "failed",
                error: { message: `provider denied ${key}` },
              },
            }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    const err = events.find((e) => e.type === "error");
    expect(err).toMatchObject({ type: "error", kind: "provider" });
    expect(err && "message" in err ? err.message : "").toContain("provider denied");
    expect(JSON.stringify(events)).not.toContain(key);
    expect(JSON.stringify(events)).toContain("[REDACTED]");
    expect(events.some((e) => e.type === "finish")).toBe(false);
  });
});
