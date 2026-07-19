import { describe, expect, it, vi } from "vitest";
import { AnthropicMessagesStreamProvider } from "../shared/provider/anthropicMessagesProvider";
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

function data(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

async function collect(
  p: AnthropicMessagesStreamProvider,
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of p.streamChat(req())) out.push(e);
  return out;
}

function make(
  fetchImpl: ReturnType<typeof vi.fn>,
  opts?: { apiKey?: string; maxTokens?: number; baseUrl?: string; model?: string },
) {
  return new AnthropicMessagesStreamProvider({
    baseUrl: opts?.baseUrl ?? "https://api.anthropic.com/v1",
    apiKey: opts?.apiKey ?? "sk-ant-secret",
    model: opts?.model ?? "claude",
    fetchImpl,
    ...(opts?.maxTokens != null ? { maxTokens: opts.maxTokens } : {}),
  });
}

describe("Anthropic Messages boundaries", () => {
  it("stop_reason mapping + empty tool args + thinking ignored", async () => {
    async function finishFor(stop: string) {
      const events = await collect(
        make(
          vi.fn(async () =>
            sseResponse([
              data({
                type: "message_start",
                message: { usage: { input_tokens: 1, output_tokens: 0 } },
              }),
              data({
                type: "message_delta",
                delta: { stop_reason: stop },
                usage: { output_tokens: 1 },
              }),
              data({ type: "message_stop" }),
              "data: [DONE]\n\n",
            ]),
          ),
        ),
      );
      return events.filter((e) => e.type !== "error").at(-1);
    }
    expect(await finishFor("end_turn")).toEqual({ type: "finish", reason: "stop" });
    expect(await finishFor("tool_use")).toEqual({ type: "finish", reason: "tool_calls" });
    expect(await finishFor("max_tokens")).toEqual({ type: "finish", reason: "length" });
    expect(await finishFor("model_context_window_exceeded")).toEqual({
      type: "finish",
      reason: "length",
    });
    expect(await finishFor("stop_sequence")).toEqual({ type: "finish", reason: "stop" });
    expect(await finishFor("refusal")).toEqual({
      type: "finish",
      reason: "content_filter",
      rawReason: "refusal",
    });
    expect(await finishFor("weird")).toEqual({
      type: "finish",
      reason: "unknown",
      rawReason: "weird",
    });

    const emptyArgs = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "message_start",
              message: { usage: { input_tokens: 1, output_tokens: 0 } },
            }),
            data({
              type: "content_block_start",
              index: 0,
              content_block: { type: "thinking", thinking: "..." },
            }),
            data({
              type: "content_block_delta",
              index: 0,
              delta: { type: "thinking_delta", thinking: "more" },
            }),
            data({ type: "content_block_stop", index: 0 }),
            data({
              type: "content_block_start",
              index: 1,
              content_block: { type: "tool_use", id: "t1", name: "host_status" },
            }),
            data({ type: "content_block_stop", index: 1 }),
            data({
              type: "message_delta",
              delta: { stop_reason: "tool_use" },
              usage: { output_tokens: 1 },
            }),
            data({ type: "message_stop" }),
            "data: [DONE]\n\n",
          ]),
        ),
      ),
    );
    expect(emptyArgs.find((e) => e.type === "tool_call_end")).toMatchObject({
      argumentsJson: "{}",
    });
    expect(emptyArgs.filter((e) => e.type === "text_delta")).toEqual([]);
  });

  it("unstopped tool / no stop_reason / malformed / missing id-name / unknown map", async () => {
    const openTool = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "content_block_start",
              index: 0,
              content_block: { type: "tool_use", id: "t", name: "host_status" },
            }),
            data({
              type: "message_delta",
              delta: { stop_reason: "tool_use" },
            }),
            data({ type: "message_stop" }),
            "data: [DONE]\n\n",
          ]),
        ),
      ),
    );
    expect(openTool.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
    expect(openTool.some((e) => e.type === "finish")).toBe(false);

    const noReason = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({ type: "message_stop" }),
            "data: [DONE]\n\n",
          ]),
        ),
      ),
    );
    expect(noReason.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const badJson = await collect(
      make(
        vi.fn(async () =>
          sseResponse(["data: {not-json\n\n", "data: [DONE]\n\n"]),
        ),
      ),
    );
    expect(badJson.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const missingId = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "content_block_start",
              index: 0,
              content_block: { type: "tool_use", name: "host_status" },
            }),
            data({ type: "message_delta", delta: { stop_reason: "end_turn" } }),
            "data: [DONE]\n\n",
          ]),
        ),
      ),
    );
    expect(missingId.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const unknown = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "content_block_start",
              index: 0,
              content_block: { type: "tool_use", id: "x", name: "not_mapped" },
            }),
            data({ type: "message_delta", delta: { stop_reason: "end_turn" } }),
            "data: [DONE]\n\n",
          ]),
        ),
      ),
    );
    expect(unknown.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);
  });

  it("HTTP/CORS/Abort/empty key/body/maxTokens; provider error redacts key", async () => {
    const emptyKey = await collect(
      make(vi.fn(), { apiKey: "   " }),
    );
    expect(emptyKey[0]).toMatchObject({ type: "error", kind: "missing_key" });

    const noFetch = vi.fn();
    const badMax = await collect(make(noFetch, { maxTokens: 0 }));
    expect(badMax[0]).toMatchObject({ type: "error", kind: "parse" });
    expect(noFetch).not.toHaveBeenCalled();

    const emptyBody = await collect(
      make(
        vi.fn(async () => new Response(null, { status: 200 })),
      ),
    );
    expect(emptyBody.some((e) => e.type === "error" && e.kind === "parse")).toBe(true);

    const http = await collect(
      make(
        vi.fn(async () =>
          new Response(JSON.stringify({ error: { message: "bad sk-ant-secret key" } }), {
            status: 401,
          }),
        ),
      ),
    );
    const httpErr = http.find((e) => e.type === "error");
    expect(httpErr).toMatchObject({ type: "error", kind: "http", status: 401 });
    expect(JSON.stringify(httpErr)).not.toContain("sk-ant-secret");
    expect(JSON.stringify(httpErr)).toContain("[REDACTED]");

    const cors = await collect(
      make(
        vi.fn(async () => {
          throw new TypeError("Failed to fetch");
        }),
      ),
    );
    expect(cors.some((e) => e.type === "error" && e.kind === "cors")).toBe(true);

    const ac = new AbortController();
    ac.abort();
    const aborted: AgentStreamEvent[] = [];
    for await (const e of make(vi.fn()).streamChat({ ...req(), signal: ac.signal })) {
      aborted.push(e);
    }
    expect(aborted[0]).toMatchObject({ type: "error", kind: "aborted" });

    const nested = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "error",
              error: { type: "api_error", message: "quota sk-ant-secret exceeded" },
            }),
          ]),
        ),
      ),
    );
    const nerr = nested.find((e) => e.type === "error");
    expect(nerr).toMatchObject({ type: "error", kind: "provider" });
    expect(JSON.stringify(nerr)).not.toContain("sk-ant-secret");
    expect(nested.some((e) => e.type === "finish")).toBe(false);
  });

  it("EOF without [DONE] still finalizes when stop_reason present", async () => {
    const events = await collect(
      make(
        vi.fn(async () =>
          sseResponse([
            data({
              type: "message_start",
              message: { usage: { input_tokens: 2, output_tokens: 0 } },
            }),
            data({
              type: "content_block_start",
              index: 0,
              content_block: { type: "text", text: "ok" },
            }),
            data({ type: "content_block_stop", index: 0 }),
            data({
              type: "message_delta",
              delta: { stop_reason: "end_turn" },
              usage: { output_tokens: 1 },
            }),
            data({ type: "message_stop" }),
          ]),
        ),
      ),
    );
    expect(events.filter((e) => e.type !== "error").at(-1)).toEqual({
      type: "finish",
      reason: "stop",
    });
  });
});
