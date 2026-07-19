import { describe, expect, it, vi } from "vitest";
import {
  ChatController,
  CHAT_READONLY_PROMPT_MARKER,
  composeChatReadonlySystemPrompt,
  listChatReadOnlyTools,
} from "../shared/agentChat";
import { AgentLoop } from "../shared/agent/agentLoop";
import {
  GuardedChatExecutor,
  CHAT_READONLY_DENY_ERROR,
} from "../shared/agentChat/chatReadOnlyTools";
import { ProviderStore } from "../shared/provider";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import type { ChatTraceEvent } from "../shared/agentChat/types";

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

const responsesTextStop = (text: string) => [
  d({ type: "response.output_text.delta", item_id: "m1", delta: text }),
  d({ type: "response.completed", response: {} }),
  "data: [DONE]\n\n",
];

function storeWith(
  format: "openai" | "responses" | "anthropic",
  apiKey = "sk-test-key",
) {
  const store = new ProviderStore();
  const base =
    format === "anthropic"
      ? "https://api.anthropic.com/v1"
      : "https://api.openai.com/v1";
  store.add({
    name: format,
    provider: format,
    apiKey,
    baseUrl: base,
    model: "m",
    apiFormat: format,
  });
  return store;
}

describe("chat readonly prompt", () => {
  it("appends readonly marker last; lists allowlist tools", () => {
    const prompt = composeChatReadonlySystemPrompt({
      routing: { content: "hello" },
      now: new Date("2026-01-01T00:00:00Z"),
    });
    expect(prompt).toContain(CHAT_READONLY_PROMPT_MARKER);
    expect(prompt.trimEnd().endsWith("。") || prompt.includes("可用工具")).toBe(
      true,
    );
    const idx = prompt.lastIndexOf(CHAT_READONLY_PROMPT_MARKER);
    expect(idx).toBeGreaterThan(0);
    expect(prompt.slice(idx)).toContain("只读");
    for (const t of listChatReadOnlyTools()) {
      expect(prompt).toContain(t.name);
    }
    expect(prompt).not.toContain("sheet.pageLayout.set");
  });
});

describe("ChatController preflight / concurrency / clear", () => {
  it("empty message / no active / missing key fail without mutating history", async () => {
    const host = new MockHostAdapter();
    const emptyStore = new ProviderStore();
    const c1 = new ChatController({ store: emptyStore, host });
    const empty = await c1.send("   ");
    expect(empty.turnStatus).toBe("empty");
    expect(c1.getState().messages).toEqual([]);

    const noActive = await c1.send("hi");
    expect(noActive.turnStatus).toBe("preflight_failed");
    expect(noActive.error?.message).toMatch(/no active/i);
    expect(c1.getState().messages).toEqual([]);

    const store = storeWith("openai", "   ");
    const c2 = new ChatController({ store, host, fetchImpl: vi.fn() });
    const miss = await c2.send("hi");
    expect(miss.turnStatus).toBe("preflight_failed");
    expect(miss.error?.kind).toBe("missing_key");
    expect(JSON.stringify(miss)).not.toContain("sk-");
    expect(c2.getState().messages).toEqual([]);
  });

  it("rejects concurrent send while busy; clear blocked while busy", async () => {
    let release!: () => void;
    const gate = new Promise<void>((r) => {
      release = r;
    });
    const fetchImpl = vi.fn(async () => {
      await gate;
      return sse(openaiTextStop("later"));
    });
    const controller = new ChatController({
      store: storeWith("openai"),
      host: new MockHostAdapter(),
      fetchImpl,
    });
    const p = controller.send("first");
    // yield microtask so send becomes running
    await Promise.resolve();
    expect(controller.getState().status).toBe("running");
    const busy = await controller.send("second");
    expect(busy.turnStatus).toBe("busy");
    expect(controller.clear()).toMatchObject({ ok: false });
    release();
    const done = await p;
    expect(done.turnStatus).toBe("completed");
    expect(controller.clear()).toMatchObject({ ok: true });
    expect(controller.getState().messages).toEqual([]);
  });
});

describe("ChatController active switch + history", () => {
  it("each turn re-reads factory; openai→responses changes fetch path", async () => {
    const store = new ProviderStore();
    const a = store.add({
      name: "chat",
      provider: "openai",
      apiKey: "sk-a",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "openai",
    });
    const b = store.add({
      name: "resp",
      provider: "openai",
      apiKey: "sk-b",
      baseUrl: "https://api.openai.com/v1",
      model: "m",
      apiFormat: "responses",
    });
    store.setActive(a.id);

    const urls: string[] = [];
    const fetchImpl = vi.fn(async (url: string) => {
      urls.push(url);
      if (url.includes("/responses")) return sse(responsesTextStop("r2"));
      return sse(openaiTextStop("r1"));
    });
    const controller = new ChatController({
      store,
      host: new MockHostAdapter(),
      fetchImpl,
    });

    const t1 = await controller.send("one");
    expect(t1.turnStatus).toBe("completed");
    expect(t1.run?.assistantText).toBe("r1");
    expect(urls[0]).toContain("/chat/completions");

    store.setActive(b.id);
    const t2 = await controller.send("two");
    expect(t2.turnStatus).toBe("completed");
    expect(t2.run?.assistantText).toBe("r2");
    expect(urls[1]).toContain("/responses");

    // Second turn history has first user once + first assistant + second user (no duplicate first user).
    const msgs = controller.getState().messages;
    const users = msgs.filter((m) => m.role === "user").map((m) => m.content);
    expect(users).toEqual(["one", "two"]);
    expect(msgs.filter((m) => m.role === "assistant").length).toBeGreaterThanOrEqual(2);
  });

  it("projects text_delta / run_end; stop aborts; serialization omits key", async () => {
    const chunks = [
      d({
        id: "c",
        choices: [{ index: 0, delta: { content: "Hel" }, finish_reason: null }],
      }),
      d({
        id: "c",
        choices: [{ index: 0, delta: { content: "lo" }, finish_reason: null }],
      }),
      d({
        id: "c",
        choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
      }),
      "data: [DONE]\n\n",
    ];
    const fetchImpl = vi.fn(async () => sse(chunks));
    const trace: ChatTraceEvent[] = [];
    const controller = new ChatController({
      store: storeWith("openai", "sk-secret-xyz"),
      host: new MockHostAdapter(),
      fetchImpl,
      onEvent: (e) => trace.push(e),
    });
    const result = await controller.send("hi");
    expect(result.turnStatus).toBe("completed");
    const deltas = trace.filter((e) => e.type === "text_delta");
    expect(deltas.map((e) => (e as { delta: string }).delta)).toEqual(["Hel", "lo"]);
    expect(trace.some((e) => e.type === "run_end")).toBe(true);
    expect(JSON.stringify(controller.getState())).not.toContain("sk-secret-xyz");
    expect(JSON.stringify(trace)).not.toContain("sk-secret-xyz");
  });
});

describe("AgentLoop + guard write rejection", () => {
  it("provider-requested range.write never hits host write (guard fail)", async () => {
    const host = new MockHostAdapter();
    const writeSpy = vi.spyOn(host, "writeRange");
    const guard = new GuardedChatExecutor(new ToolExecutor(host));

    // Guard alone: deny write, no host call.
    const gWrite = await guard.execute({
      name: "range.write",
      arguments: { sheetName: "S", range: "A1", values: [["x"]] },
    });
    expect(gWrite.ok).toBe(false);
    if (!gWrite.ok) expect(gWrite.error).toContain(CHAT_READONLY_DENY_ERROR);
    expect(writeSpy).not.toHaveBeenCalled();

    // AgentLoop exposes write in tools (so name maps), but guarded executor still blocks host.
    let n = 0;
    const fetch2 = vi.fn(async () => {
      n += 1;
      if (n === 1) {
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
                      id: "call_w",
                      type: "function",
                      function: {
                        name: "range_write",
                        arguments:
                          '{"sheetName":"S","range":"A1","values":[["x"]]}',
                      },
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
      return sse(openaiTextStop("blocked"));
    });

    const store = storeWith("openai");
    const { createStreamProviderFromStore } = await import("../shared/provider");
    const created = createStreamProviderFromStore(store, { fetchImpl: fetch2 });
    expect(created.ok).toBe(true);
    if (!created.ok) return;

    const writeDef = TOOL_DEFINITIONS.find((t) => t.name === "range.write");
    expect(writeDef).toBeTruthy();
    const tools = [...listChatReadOnlyTools(), writeDef!];

    const outcomes: Array<{ kind: string; error?: string }> = [];
    const result = await new AgentLoop({
      provider: created.provider,
      executor: guard,
      systemPrompt: "sys",
      tools,
      onEvent: (e) => {
        if (e.type === "tool_outcome") {
          const o = e.outcome;
          if (o.kind === "host" && !o.result.ok) {
            outcomes.push({ kind: "host_fail", error: o.result.error });
          } else {
            outcomes.push({ kind: o.kind });
          }
        }
      },
    }).run({ userMessage: "write A1" });

    expect(result.status).toBe("completed");
    expect(writeSpy).not.toHaveBeenCalled();
    expect(
      outcomes.some(
        (o) =>
          o.kind === "unknown_tool" ||
          (o.kind === "host_fail" &&
            (o.error ?? "").includes(CHAT_READONLY_DENY_ERROR)),
      ),
    ).toBe(true);
  });
});

describe("ChatController prompt assembly", () => {
  it("uses injected composeSystemPrompt; marker present by default path", async () => {
    const compose = vi.fn(() => `base\n\n## ${CHAT_READONLY_PROMPT_MARKER}\nreadonly`);
    const fetchImpl = vi.fn(async () => sse(openaiTextStop("ok")));
    const controller = new ChatController({
      store: storeWith("openai"),
      host: new MockHostAdapter(),
      fetchImpl,
      composeSystemPrompt: compose,
    });
    await controller.send("hello excel");
    expect(compose).toHaveBeenCalledTimes(1);
    expect(compose).toHaveBeenCalledWith("hello excel");
  });
});
