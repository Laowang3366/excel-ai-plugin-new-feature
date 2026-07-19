import { describe, expect, it, vi } from "vitest";
import { AgentLoop } from "../shared/agent/agentLoop";
import { OpenAIResponsesStreamProvider } from "../shared/provider/openaiResponsesProvider";
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

describe("OpenAI Responses + AgentLoop e2e", () => {
  it("round1 host.status via call_id then round2 final text; second input matches", async () => {
    const bodies: Record<string, unknown>[] = [];
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      bodies.push(JSON.parse(String(init?.body)));
      if (fetchImpl.mock.calls.length === 1) {
        return sse([
          d({
            type: "response.output_item.added",
            item: {
              type: "function_call",
              id: "fc_1",
              call_id: "call_1",
              name: "host_status",
            },
          }),
          d({
            type: "response.function_call_arguments.done",
            item_id: "fc_1",
            arguments: "{}",
          }),
          d({
            type: "response.output_item.done",
            item: {
              type: "function_call",
              id: "fc_1",
              call_id: "call_1",
              name: "host_status",
              arguments: "{}",
            },
          }),
          d({ type: "response.completed", response: {} }),
          "data: [DONE]\n\n",
        ]);
      }
      return sse([
        d({
          type: "response.output_text.delta",
          item_id: "m1",
          delta: "All good",
        }),
        d({ type: "response.completed", response: {} }),
        "data: [DONE]\n\n",
      ]);
    });

    const execute = vi.spyOn(ToolExecutor.prototype, "execute");
    const result = await new AgentLoop({
      provider: new OpenAIResponsesStreamProvider({
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk-test",
        model: "gpt-4o",
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

    const second = bodies[1] as { input: Array<Record<string, unknown>> };
    expect(second.input).toEqual(
      expect.arrayContaining([
        { type: "message", role: "user", content: "check host" },
        expect.objectContaining({
          type: "function_call",
          call_id: "call_1",
          name: "host_status",
        }),
        expect.objectContaining({
          type: "function_call_output",
          call_id: "call_1",
        }),
      ]),
    );
    const fc = second.input.find((x) => x.type === "function_call") as {
      call_id: string;
    };
    const out = second.input.find((x) => x.type === "function_call_output") as {
      call_id: string;
    };
    expect(fc.call_id).toBe("call_1");
    expect(out.call_id).toBe(fc.call_id);
    execute.mockRestore();
  });
});
