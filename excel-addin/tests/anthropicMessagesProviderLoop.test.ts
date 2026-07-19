import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import { AnthropicMessagesStreamProvider } from "../shared/provider/anthropicMessagesProvider";
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

describe("AnthropicMessagesStreamProvider + AgentLoop", () => {
  it("two rounds: host.status then final text; second body merges tool_result with same id", async () => {
    const bodies: unknown[] = [];
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(JSON.parse(String(init?.body)));
      if (fetchImpl.mock.calls.length === 1) {
        return sse([
          d({
            type: "message_start",
            message: { usage: { input_tokens: 3, output_tokens: 0 } },
          }),
          d({
            type: "content_block_start",
            index: 0,
            content_block: { type: "tool_use", id: "toolu_1", name: "host_status" },
          }),
          d({
            type: "content_block_delta",
            index: 0,
            delta: { type: "input_json_delta", partial_json: "{}" },
          }),
          d({ type: "content_block_stop", index: 0 }),
          d({
            type: "message_delta",
            delta: { stop_reason: "tool_use" },
            usage: { output_tokens: 2 },
          }),
          d({ type: "message_stop" }),
          "data: [DONE]\n\n",
        ]);
      }
      return sse([
        d({
          type: "message_start",
          message: { usage: { input_tokens: 8, output_tokens: 0 } },
        }),
        d({
          type: "content_block_start",
          index: 0,
          content_block: { type: "text", text: "" },
        }),
        d({
          type: "content_block_delta",
          index: 0,
          delta: { type: "text_delta", text: "All good" },
        }),
        d({ type: "content_block_stop", index: 0 }),
        d({
          type: "message_delta",
          delta: { stop_reason: "end_turn" },
          usage: { output_tokens: 2 },
        }),
        d({ type: "message_stop" }),
        "data: [DONE]\n\n",
      ]);
    });

    const execute = vi.spyOn(ToolExecutor.prototype, "execute");
    const result = await new AgentLoop({
      provider: new AnthropicMessagesStreamProvider({
        baseUrl: "https://api.anthropic.com/v1",
        apiKey: "sk-ant-test",
        model: "claude-test",
        fetchImpl,
      }),
      executor: new ToolExecutor(new MockHostAdapter()),
      systemPrompt: "sys",
      tools: TOOL_DEFINITIONS.filter((t) => t.name === "host.status"),
    }).run({ userMessage: "check host" });

    expect(result.status).toBe("completed");
    expect(result.rounds).toBe(2);
    expect(result.assistantText).toBe("All good");
    expect(execute).toHaveBeenCalledTimes(1);
    expect(execute.mock.calls[0]?.[0]).toMatchObject({
      name: "host.status",
      arguments: {},
    });
    expect(fetchImpl).toHaveBeenCalledTimes(2);

    const first = bodies[0] as {
      system?: string;
      messages: Array<Record<string, unknown>>;
      tools?: unknown[];
    };
    expect(first.system).toBe("sys");
    expect(first.messages).toEqual([{ role: "user", content: "check host" }]);

    const second = bodies[1] as {
      system?: string;
      messages: Array<Record<string, unknown>>;
    };
    expect(second.system).toBe("sys");
    // assistant tool_use + merged tool_result user with same id
    const assistant = second.messages.find(
      (m) => m.role === "assistant" && Array.isArray(m.content),
    ) as { content: Array<Record<string, unknown>> };
    expect(assistant).toBeTruthy();
    const toolUse = assistant.content.find((b) => b.type === "tool_use") as {
      id: string;
      name: string;
    };
    expect(toolUse).toMatchObject({ id: "toolu_1", name: "host_status" });

    const toolUser = second.messages.find(
      (m) =>
        m.role === "user" &&
        Array.isArray(m.content) &&
        (m.content as Array<Record<string, unknown>>).some(
          (b) => b.type === "tool_result",
        ),
    ) as { content: Array<Record<string, unknown>> };
    expect(toolUser).toBeTruthy();
    const toolResult = toolUser.content.find((b) => b.type === "tool_result") as {
      tool_use_id: string;
    };
    expect(toolResult.tool_use_id).toBe(toolUse.id);
    expect(toolResult.tool_use_id).toBe("toolu_1");
    execute.mockRestore();
  });
});
