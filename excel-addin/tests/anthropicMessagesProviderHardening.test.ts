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

function data(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

type ReaderHooks = {
  releaseLock: ReturnType<typeof vi.fn>;
  readSeq: Array<{ done: boolean; value?: Uint8Array }>;
};

/** Response body whose getReader is instrumented for releaseLock. */
function instrumentedResponse(
  chunks: string[],
  hooks?: Partial<ReaderHooks>,
): { response: Response; releaseLock: ReturnType<typeof vi.fn> } {
  const enc = new TextEncoder();
  const encoded = chunks.map((c) => enc.encode(c));
  const releaseLock = hooks?.releaseLock ?? vi.fn();
  const readSeq =
    hooks?.readSeq ??
    [
      ...encoded.map((value) => ({ done: false as const, value })),
      { done: true as const },
    ];
  let readIndex = 0;
  const reader = {
    read: vi.fn(async () => {
      const next = readSeq[readIndex] ?? { done: true as const };
      readIndex += 1;
      return next;
    }),
    releaseLock,
    cancel: vi.fn(),
  };
  // Provider only needs ok + body.getReader(); avoid Response wrapping which may ignore custom body.
  const response = {
    ok: true,
    status: 200,
    body: { getReader: () => reader },
  } as unknown as Response;
  return { response, releaseLock };
}

async function collect(
  provider: AnthropicMessagesStreamProvider,
  request: StreamChatRequest = req(),
): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(request)) out.push(e);
  return out;
}

function provider(fetchImpl: ReturnType<typeof vi.fn>) {
  return new AnthropicMessagesStreamProvider({
    baseUrl: "https://api.anthropic.com/v1",
    apiKey: "sk-ant-secret",
    model: "claude",
    fetchImpl,
  });
}

const happyChunks = [
  data({
    type: "message_start",
    message: { usage: { input_tokens: 3, output_tokens: 0 } },
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
  "data: [DONE]\n\n",
];

describe("AnthropicMessagesStreamProvider hardening", () => {
  it("releaseLock on normal completion (including [DONE] early stop)", async () => {
    const { response, releaseLock } = instrumentedResponse(happyChunks);
    const fetchImpl = vi.fn(async () => response);
    const events = await collect(provider(fetchImpl));
    expect(events.some((e) => e.type === "finish" && e.reason === "stop")).toBe(
      true,
    );
    expect(releaseLock).toHaveBeenCalledTimes(1);
  });

  it("releaseLock on parse error early return (malformed SSE)", async () => {
    const { response, releaseLock } = instrumentedResponse([
      "data: {not-json\n\n",
      "data: [DONE]\n\n",
    ]);
    const fetchImpl = vi.fn(async () => response);
    const events = await collect(provider(fetchImpl));
    expect(events.some((e) => e.type === "error" && e.kind === "parse")).toBe(
      true,
    );
    expect(events.some((e) => e.type === "finish")).toBe(false);
    expect(releaseLock).toHaveBeenCalledTimes(1);
  });

  it("releaseLock on provider error event early return", async () => {
    const { response, releaseLock } = instrumentedResponse([
      data({
        type: "error",
        error: { type: "api_error", message: "boom" },
      }),
    ]);
    const fetchImpl = vi.fn(async () => response);
    const events = await collect(provider(fetchImpl));
    expect(events.some((e) => e.type === "error" && e.kind === "provider")).toBe(
      true,
    );
    expect(releaseLock).toHaveBeenCalledTimes(1);
  });

  it("releaseLock still called when releaseLock itself throws", async () => {
    const releaseLock = vi.fn(() => {
      throw new Error("release failed");
    });
    const { response } = instrumentedResponse(happyChunks, { releaseLock });
    const fetchImpl = vi.fn(async () => response);
    const events = await collect(provider(fetchImpl));
    // Stream result must not be masked by releaseLock throw.
    expect(events.some((e) => e.type === "finish" && e.reason === "stop")).toBe(
      true,
    );
    expect(releaseLock).toHaveBeenCalledTimes(1);
  });

  it("cache_creation-only does not set cachedInputTokens; cache_read does", async () => {
    const creationOnly = [
      data({
        type: "message_start",
        message: {
          usage: {
            input_tokens: 10,
            output_tokens: 0,
            cache_creation_input_tokens: 9,
          },
        },
      }),
      data({
        type: "message_delta",
        delta: { stop_reason: "end_turn" },
        usage: { output_tokens: 2 },
      }),
      data({ type: "message_stop" }),
      "data: [DONE]\n\n",
    ];
    const { response: r1 } = instrumentedResponse(creationOnly);
    const events1 = await collect(provider(vi.fn(async () => r1)));
    const usages1 = events1.filter((e) => e.type === "usage") as Array<{
      type: "usage";
      usage: { inputTokens: number; outputTokens: number; cachedInputTokens?: number };
    }>;
    expect(usages1.length).toBeGreaterThanOrEqual(1);
    for (const u of usages1) {
      expect(u.usage.cachedInputTokens).toBeUndefined();
    }
    expect(usages1.at(-1)!.usage).toMatchObject({
      inputTokens: 10,
      outputTokens: 2,
    });

    const withRead = [
      data({
        type: "message_start",
        message: {
          usage: {
            input_tokens: 12,
            output_tokens: 0,
            cache_read_input_tokens: 4,
            cache_creation_input_tokens: 99,
          },
        },
      }),
      data({
        type: "message_delta",
        delta: { stop_reason: "end_turn" },
        usage: { output_tokens: 3 },
      }),
      data({ type: "message_stop" }),
      "data: [DONE]\n\n",
    ];
    const { response: r2 } = instrumentedResponse(withRead);
    const events2 = await collect(provider(vi.fn(async () => r2)));
    const usages2 = events2.filter((e) => e.type === "usage") as Array<{
      type: "usage";
      usage: { inputTokens: number; outputTokens: number; cachedInputTokens?: number };
    }>;
    const last = usages2.at(-1)!;
    expect(last.usage.cachedInputTokens).toBe(4);
    expect(last.usage.inputTokens).toBe(12);
    expect(last.usage.outputTokens).toBe(3);
  });
});
