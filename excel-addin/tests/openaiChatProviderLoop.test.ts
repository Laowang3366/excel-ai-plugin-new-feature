import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import { OpenAIChatCompletionsStreamProvider } from "../shared/provider/openaiChatCompletionsProvider";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

function sseBody(chunks: string[]): Response {
  const encoder = new TextEncoder();
  let i = 0;
  return new Response(
    new ReadableStream<Uint8Array>({
      pull(controller) {
        if (i >= chunks.length) {
          controller.close();
          return;
        }
        controller.enqueue(encoder.encode(chunks[i++]));
      },
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } },
  );
}

describe("OpenAI provider + AgentLoop e2e (fake SSE)", () => {
  it("round1 host.status tool then round2 final text", async () => {
    const bodies: Record<string, unknown>[] = [];
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(JSON.parse(String(init?.body)));
      if (fetchImpl.mock.calls.length === 1) {
        return sseBody([
          'data: {"choices":[{"delta":{"tool_calls":[{"index":0,"id":"call_1","function":{"name":"host_status","arguments":"{}"}}]},"finish_reason":"tool_calls"}]}\n\n',
          "data: [DONE]\n\n",
        ]);
      }
      return sseBody([
        'data: {"choices":[{"delta":{"content":"All good"},"finish_reason":"stop"}]}\n\n',
        "data: [DONE]\n\n",
      ]);
    });

    const host = new MockHostAdapter();
    const execute = vi.spyOn(ToolExecutor.prototype, "execute");
    const executor = new ToolExecutor(host);
    const provider = new OpenAIChatCompletionsStreamProvider({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-test",
      model: "gpt-4o",
      fetchImpl,
    });
    const subset = TOOL_DEFINITIONS.filter((t) => t.name === "host.status");
    const result = await new AgentLoop({
      provider,
      executor,
      systemPrompt: "sys",
      tools: subset,
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

    const second = bodies[1] as {
      messages: Array<Record<string, unknown>>;
    };
    const roles = second.messages.map((m) => m.role);
    expect(roles).toEqual(["system", "user", "assistant", "tool"]);
    const assistant = second.messages[2] as {
      tool_calls: Array<{ function: { name: string }; id: string }>;
    };
    expect(assistant.tool_calls[0]?.id).toBe("call_1");
    expect(assistant.tool_calls[0]?.function.name).toBe("host_status");
    const toolMsg = second.messages[3] as { tool_call_id: string; content: string };
    expect(toolMsg.tool_call_id).toBe("call_1");
    expect(toolMsg.tool_call_id).toBe(assistant.tool_calls[0]?.id);
    expect(toolMsg.content).toContain('"kind":"host"');
    execute.mockRestore();
  });
});
