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
  return `data: ${JSON.stringify(obj)}\n\n`;
}

describe("OpenAI Responses error boundaries", () => {
  it("missing call_id / unknown name / no terminal / failed / HTTP key redact / empty key", async () => {
    const missingCall = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({
              type: "response.output_item.done",
              item: { type: "function_call", id: "i", name: "host_status", arguments: "{}" },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(missingCall.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const unknown = await collect(
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
                id: "i",
                call_id: "c",
                name: "not_mapped",
                arguments: "{}",
              },
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(unknown.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const noTerm = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({ type: "response.output_text.delta", item_id: "m", delta: "x" }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(noTerm.at(-1)).toMatchObject({ type: "error", kind: "parse" });

    const failed = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({ type: "response.failed", error: { message: "boom" } }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(failed.some((e) => e.type === "error" && e.kind === "provider")).toBe(true);
    expect(failed.some((e) => e.type === "finish")).toBe(false);

    const http = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk-secret",
        model: "m",
        fetchImpl: vi.fn(
          async () =>
            new Response(JSON.stringify({ error: { message: "nope sk-secret" } }), {
              status: 401,
            }),
        ),
      }),
    );
    expect(http[0]).toMatchObject({ type: "error", kind: "http", status: 401 });
    expect(JSON.stringify(http)).not.toContain("sk-secret");

    const noKey = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "",
        model: "m",
        fetchImpl: vi.fn(),
      }),
    );
    expect(noKey[0]).toMatchObject({ type: "error", kind: "missing_key" });
  });

  it("CORS / abort / empty body / malformed / reasoning ignored", async () => {
    const cors = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => {
          throw new TypeError("Failed to fetch");
        }),
      }),
    );
    expect(cors[0]).toMatchObject({ type: "error", kind: "cors" });

    const aborted = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => {
          throw Object.assign(new Error("aborted"), { name: "AbortError" });
        }),
      }),
    );
    expect(aborted[0]).toMatchObject({ type: "error", kind: "aborted" });

    const empty = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(
          async () =>
            new Response(null, {
              status: 200,
              headers: { "Content-Type": "text/event-stream" },
            }),
        ),
      }),
    );
    expect(empty[0]).toMatchObject({ type: "error", kind: "parse" });

    const malformed = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () => sseResponse(["data: {bad}\n\n"])),
      }),
    );
    expect(malformed.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const events = await collect(
      new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "m",
        fetchImpl: vi.fn(async () =>
          sseResponse([
            data({ type: "response.reasoning.delta", delta: "think" }),
            data({
              type: "response.output_text.delta",
              item_id: "m",
              delta: "ok",
            }),
            data({ type: "response.completed", response: {} }),
            "data: [DONE]\n\n",
          ]),
        ),
      }),
    );
    expect(events.filter((e) => e.type === "text_delta")).toEqual([
      { type: "text_delta", delta: "ok" },
    ]);
    expect(events.some((e) => e.type === "error")).toBe(false);
  });
});
