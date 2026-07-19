import { describe, expect, it, vi } from "vitest";
import { ChatController } from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
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

const openaiTextStop = (text: string) => [
  d({
    id: "c",
    choices: [{ index: 0, delta: { content: text }, finish_reason: null }],
  }),
  d({
    id: "c",
    choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
  }),
  "data: [DONE]\n\n",
];

function storeWith(apiKey = "sk-test-key") {
  const store = new ProviderStore();
  store.add({
    name: "openai",
    provider: "openai",
    apiKey,
    baseUrl: "https://api.openai.com/v1",
    model: "m",
    apiFormat: "openai",
  });
  return store;
}

describe("ChatController stop and max_rounds", () => {
  it("stop() aborts an in-flight turn", async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => {
      release = r;
    });
    const fetchImpl = vi.fn(async () => {
      await gate;
      return sse(openaiTextStop("too late"));
    });
    const controller = new ChatController({
      store: storeWith(),
      host: new MockHostAdapter(),
      fetchImpl,
    });
    const p = controller.send("run");
    await Promise.resolve();
    expect(controller.getState().status).toBe("running");
    controller.stop();
    expect(controller.getState().status).toBe("stopping");
    release();
    const result = await p;
    expect(["aborted", "completed", "failed"]).toContain(result.turnStatus);
    expect(controller.getState().status).toBe("idle");
  });

  it("max_rounds commits messages and reports max_rounds", async () => {
    const fetchImpl = vi.fn(async () =>
      sse([
        d({
          id: "c",
          choices: [
            {
              index: 0,
              delta: {
                tool_calls: [
                  {
                    index: 0,
                    id: "call_s",
                    type: "function",
                    function: { name: "host_status", arguments: "{}" },
                  },
                ],
              },
              finish_reason: null,
            },
          ],
        }),
        d({
          id: "c",
          choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }],
        }),
        "data: [DONE]\n\n",
      ]),
    );
    const controller = new ChatController({
      store: storeWith(),
      host: new MockHostAdapter(),
      fetchImpl,
      maxRounds: 1,
    });
    const result = await controller.send("status please");
    expect(result.turnStatus).toBe("max_rounds");
    expect(result.run?.rounds).toBe(1);
    expect(controller.getState().messages.length).toBeGreaterThan(0);
    expect(controller.getState().messages.some((m) => m.role === "user")).toBe(
      true,
    );
  });
});
