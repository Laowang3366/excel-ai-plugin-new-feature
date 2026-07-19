import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  createStreamProvider,
  createStreamProviderFromStore,
  ProviderStore,
} from "../shared/provider";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

function sse(chunks: string[]): Response {
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

function d(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

describe("createStreamProvider + AgentLoop smoke", () => {
  it("factory openai provider drives real AgentLoop tool round-trip", async () => {
    const bodies: unknown[] = [];
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(JSON.parse(String(init?.body)));
      if (fetchImpl.mock.calls.length === 1) {
        return sse([
          d({
            id: "c1",
            choices: [
              {
                index: 0,
                delta: {
                  tool_calls: [
                    {
                      index: 0,
                      id: "call_1",
                      type: "function",
                      function: { name: "host_status", arguments: "" },
                    },
                  ],
                },
                finish_reason: null,
              },
            ],
          }),
          d({
            id: "c1",
            choices: [
              {
                index: 0,
                delta: {
                  tool_calls: [
                    {
                      index: 0,
                      function: { arguments: "{}" },
                    },
                  ],
                },
                finish_reason: null,
              },
            ],
          }),
          d({
            id: "c1",
            choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }],
          }),
          "data: [DONE]\n\n",
        ]);
      }
      return sse([
        d({
          id: "c2",
          choices: [{ index: 0, delta: { content: "Host ready" }, finish_reason: null }],
        }),
        d({
          id: "c2",
          choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
        }),
        "data: [DONE]\n\n",
      ]);
    });

    const created = createStreamProvider({
      apiFormat: "openai",
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-loop",
      model: "gpt-4o",
      fetchImpl,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;

    const execute = vi.spyOn(ToolExecutor.prototype, "execute");
    const result = await new AgentLoop({
      provider: created.provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "sys",
      tools: TOOL_DEFINITIONS.filter((t) => t.name === "host.status"),
    }).run({ userMessage: "status?" });

    expect(result.status).toBe("completed");
    expect(result.rounds).toBe(2);
    expect(result.assistantText).toBe("Host ready");
    expect(execute).toHaveBeenCalledTimes(1);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    expect((bodies[0] as { model: string }).model).toBe("gpt-4o");
    // second request includes tool result linkage
    const second = bodies[1] as { messages: Array<Record<string, unknown>> };
    expect(second.messages.some((m) => m.role === "tool")).toBe(true);
    expect(JSON.stringify(bodies)).not.toContain("sk-loop");
    execute.mockRestore();
  });

  it("store-created anthropic provider also works with AgentLoop", async () => {
    const fetchImpl = vi.fn(async () =>
      sse([
        d({
          type: "message_start",
          message: { usage: { input_tokens: 1, output_tokens: 0 } },
        }),
        d({
          type: "content_block_start",
          index: 0,
          content_block: { type: "text", text: "done" },
        }),
        d({ type: "content_block_stop", index: 0 }),
        d({
          type: "message_delta",
          delta: { stop_reason: "end_turn" },
          usage: { output_tokens: 1 },
        }),
        d({ type: "message_stop" }),
        "data: [DONE]\n\n",
      ]),
    );
    const store = new ProviderStore();
    store.add({
      name: "Claude",
      provider: "anthropic",
      apiKey: "sk-ant-loop",
      baseUrl: "https://api.anthropic.com/v1",
      model: "claude",
      apiFormat: "anthropic",
    });
    const created = createStreamProviderFromStore(store, { fetchImpl });
    expect(created.ok).toBe(true);
    if (!created.ok) return;

    const result = await new AgentLoop({
      provider: created.provider,
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "sys",
      tools: TOOL_DEFINITIONS.filter((t) => t.name === "host.status"),
    }).run({ userMessage: "hi" });

    expect(result.status).toBe("completed");
    expect(result.assistantText).toBe("done");
    const call0 = fetchImpl.mock.calls[0] as unknown as [string, RequestInit];
    expect(call0[0]).toBe("https://api.anthropic.com/v1/messages");
  });
});
