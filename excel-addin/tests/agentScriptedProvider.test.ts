import { describe, expect, it, vi } from "vitest";
import {
  ScriptedStreamProvider,
  errorEvent,
  textThenStop,
  toolCallThenFinish,
} from "../shared/agent/scriptedProvider";
import type { StreamChatRequest } from "../shared/agent/types";
import type { ToolDefinition } from "../shared/tools/types";

const tools: ToolDefinition[] = [
  {
    name: "host.status",
    description: "d",
    riskLevel: "safe",
    parameters: { type: "object", properties: {}, additionalProperties: false },
  },
];

function baseReq(signal?: AbortSignal): StreamChatRequest {
  return {
    systemPrompt: "sys",
    messages: [{ role: "user", content: "hi" }],
    tools,
    signal,
  };
}

async function collect(provider: ScriptedStreamProvider, req: StreamChatRequest) {
  const out = [];
  for await (const e of provider.streamChat(req)) out.push(e);
  return out;
}

describe("ScriptedStreamProvider", () => {
  it("advances rounds and supports function scripts with callCount", async () => {
    const provider = new ScriptedStreamProvider({
      rounds: [
        textThenStop("one"),
        ({ callCount }) => textThenStop(`two-${callCount}`),
      ],
    });
    expect(await collect(provider, baseReq())).toEqual(textThenStop("one"));
    expect(provider.callCount).toBe(1);
    expect(await collect(provider, baseReq())).toEqual(textThenStop("two-2"));
    expect(provider.callCount).toBe(2);
  });

  it("stores shallow-copied lastRequest (messages/tools sliced)", async () => {
    const provider = new ScriptedStreamProvider({ rounds: [textThenStop("x")] });
    const messages = [{ role: "user" as const, content: "hi" }];
    const reqTools = tools.slice();
    const req = { systemPrompt: "s", messages, tools: reqTools };
    await collect(provider, req);
    expect(provider.lastRequest?.messages).toEqual(messages);
    expect(provider.lastRequest?.messages).not.toBe(messages);
    expect(provider.lastRequest?.tools).not.toBe(reqTools);
    // element identity is shallow
    expect(provider.lastRequest?.messages[0]).toBe(messages[0]);
  });

  it("exhausted error emits provider error without finish; stop emits finish", async () => {
    const errP = new ScriptedStreamProvider({
      rounds: [textThenStop("only")],
      onExhausted: "error",
    });
    await collect(errP, baseReq());
    const exhausted = await collect(errP, baseReq());
    expect(exhausted).toEqual([
      { type: "error", message: "Scripted provider exhausted", kind: "provider" },
    ]);

    const stopP = new ScriptedStreamProvider({
      rounds: [],
      onExhausted: "stop",
    });
    expect(await collect(stopP, baseReq())).toEqual([{ type: "finish", reason: "stop" }]);
  });

  it("aborts at start and between events", async () => {
    const start = new AbortController();
    start.abort();
    const provider = new ScriptedStreamProvider({ rounds: [textThenStop("x")] });
    await expect(collect(provider, baseReq(start.signal))).rejects.toMatchObject({
      name: "AbortError",
    });

    const mid = new AbortController();
    const delayed = new ScriptedStreamProvider({
      rounds: [
        [
          { type: "text_delta", delta: "a" },
          { type: "text_delta", delta: "b" },
          { type: "finish", reason: "stop" },
        ],
      ],
      eventDelayMs: 30,
    });
    const iter = delayed.streamChat(baseReq(mid.signal))[Symbol.asyncIterator]();
    const first = await iter.next();
    expect(first.value).toEqual({ type: "text_delta", delta: "a" });
    mid.abort();
    await expect(iter.next()).rejects.toMatchObject({ name: "AbortError" });
  });

  it("never calls fetch; factories produce expected events", async () => {
    const fetchSpy = vi.spyOn(globalThis, "fetch");
    const provider = new ScriptedStreamProvider({
      rounds: [
        toolCallThenFinish("1", "host.status", "{}"),
        errorEvent("nope", "http", 401),
      ],
    });
    expect(await collect(provider, baseReq())).toEqual(
      toolCallThenFinish("1", "host.status", "{}"),
    );
    expect(await collect(provider, baseReq())).toEqual([
      { type: "error", message: "nope", kind: "http", status: 401 },
    ]);
    expect(fetchSpy).not.toHaveBeenCalled();
    fetchSpy.mockRestore();
  });
});
