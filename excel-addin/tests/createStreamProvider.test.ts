import { describe, expect, it, vi } from "vitest";
import {
  createStreamProvider,
  createStreamProviderFromStore,
  ProviderStore,
} from "../shared/provider";
import type {
  AgentStreamEvent,
  StreamChatRequest,
} from "../shared/agent/types";
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

async function drain(provider: {
  streamChat: (r: StreamChatRequest) => AsyncIterable<AgentStreamEvent>;
}): Promise<AgentStreamEvent[]> {
  const out: AgentStreamEvent[] = [];
  for await (const e of provider.streamChat(req())) out.push(e);
  return out;
}

const openaiDone = [
  data({
    id: "c",
    choices: [{ index: 0, delta: { content: "ok" }, finish_reason: null }],
  }),
  data({ id: "c", choices: [{ index: 0, delta: {}, finish_reason: "stop" }] }),
  "data: [DONE]\n\n",
];

const responsesDone = [
  data({ type: "response.output_text.delta", item_id: "m1", delta: "ok" }),
  data({ type: "response.completed", response: {} }),
  "data: [DONE]\n\n",
];

const anthropicDone = [
  data({
    type: "message_start",
    message: { usage: { input_tokens: 1, output_tokens: 0 } },
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

describe("createStreamProvider routing", () => {
  it("openai apiFormat hits /chat/completions with Bearer", async () => {
    const fetchImpl = vi.fn(async () => sseResponse(openaiDone));
    const created = createStreamProvider({
      apiFormat: "openai",
      provider: "openai",
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-openai-secret",
      model: "gpt-4o",
      reasoningMode: "max",
      fetchImpl,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    const events = await drain(created.provider);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [
      string,
      RequestInit,
    ];
    expect(url).toBe("https://api.openai.com/v1/chat/completions");
    const headers = init.headers as Record<string, string>;
    expect(headers.Authorization).toBe("Bearer sk-openai-secret");
    const body = JSON.parse(String(init.body));
    expect(body.reasoning).toEqual({ effort: "xhigh" });
    expect(body).not.toHaveProperty("reasoning_effort");
    expect(JSON.stringify(body)).not.toContain("sk-openai-secret");
    expect(events.some((e) => e.type === "finish")).toBe(true);
  });

  it("responses apiFormat hits /responses with Bearer", async () => {
    const fetchImpl = vi.fn(async () => sseResponse(responsesDone));
    const created = createStreamProvider({
      apiFormat: "responses",
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-resp-secret",
      model: "gpt-4o",
      reasoningMode: "max",
      fetchImpl,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    await drain(created.provider);
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [
      string,
      RequestInit,
    ];
    expect(url).toBe("https://api.openai.com/v1/responses");
    expect((init.headers as Record<string, string>).Authorization).toBe(
      "Bearer sk-resp-secret",
    );
    const body = JSON.parse(String(init.body));
    expect(body.reasoning).toEqual({ effort: "xhigh", summary: "auto" });
    expect(JSON.stringify(body)).not.toContain("sk-resp-secret");
  });

  it("anthropic apiFormat hits /messages with x-api-key (no Bearer)", async () => {
    const fetchImpl = vi.fn(async () => sseResponse(anthropicDone));
    const created = createStreamProvider({
      apiFormat: "anthropic",
      baseUrl: "https://api.anthropic.com/v1",
      apiKey: "sk-ant-secret",
      model: "claude-3-5-sonnet-latest",
      reasoningMode: "max",
      fetchImpl,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    await drain(created.provider);
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [
      string,
      RequestInit,
    ];
    expect(url).toBe("https://api.anthropic.com/v1/messages");
    const headers = init.headers as Record<string, string>;
    expect(headers["x-api-key"]).toBe("sk-ant-secret");
    expect(headers.Authorization).toBeUndefined();
    expect(headers["anthropic-version"]).toBe("2023-06-01");
    const body = JSON.parse(String(init.body));
    expect(body.thinking).toEqual({ type: "adaptive" });
    expect(body.output_config).toEqual({ effort: "xhigh" });
    expect(JSON.stringify(body)).not.toContain("sk-ant-secret");
  });

  for (const testCase of [
    { apiFormat: "openai", response: openaiDone },
    { apiFormat: "anthropic", response: anthropicDone },
  ] as const) {
    it(`${testCase.apiFormat} off/invalid reasoning leaves the direct request unchanged`, async () => {
      for (const reasoningMode of ["off", "invalid"] as const) {
        let body: Record<string, unknown> | undefined;
        const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
          body = JSON.parse(String(init?.body));
          return sseResponse([...testCase.response]);
        });
        const created = createStreamProvider({
          apiFormat: testCase.apiFormat,
          provider: testCase.apiFormat === "openai" ? "openai" : "anthropic",
          baseUrl:
            testCase.apiFormat === "anthropic"
              ? "https://api.anthropic.com/v1"
              : "https://api.openai.com/v1",
          apiKey: "sk-off-secret",
          model: "test-model",
          reasoningMode: reasoningMode as never,
          fetchImpl,
        });
        expect(created.ok).toBe(true);
        if (!created.ok) continue;
        await drain(created.provider);
        expect(body).not.toHaveProperty("reasoning_effort");
        expect(body).not.toHaveProperty("reasoning");
        expect(body).not.toHaveProperty("thinking");
        expect(body).not.toHaveProperty("output_config");
        expect(JSON.stringify(body)).not.toContain("sk-off-secret");
      }
    });
  }

  it("responses off sends explicit effort=none; invalid omits reasoning", async () => {
    for (const [reasoningMode, expected] of [
      ["off", { effort: "none" }],
      ["invalid", undefined],
    ] as const) {
      let body: Record<string, unknown> | undefined;
      const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
        body = JSON.parse(String(init?.body));
        return sseResponse([...responsesDone]);
      });
      const created = createStreamProvider({
        apiFormat: "responses",
        provider: "openai",
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk-off-secret",
        model: "test-model",
        reasoningMode: reasoningMode as never,
        fetchImpl,
      });
      expect(created.ok).toBe(true);
      if (!created.ok) continue;
      await drain(created.provider);
      if (expected) {
        expect(body?.reasoning).toEqual(expected);
      } else {
        expect(body).not.toHaveProperty("reasoning");
      }
      expect(body).not.toHaveProperty("reasoning_effort");
      expect(JSON.stringify(body)).not.toContain("sk-off-secret");
    }
  });

  it("unknown/empty format and empty fields fail without fetch; errors omit key", () => {
    const fetchImpl = vi.fn();
    const key = "sk-must-not-leak";
    const cases = [
      createStreamProvider({
        apiFormat: "weird",
        baseUrl: "https://x",
        apiKey: key,
        model: "m",
        fetchImpl,
      }),
      createStreamProvider({
        apiFormat: "  ",
        baseUrl: "https://x",
        apiKey: key,
        model: "m",
        fetchImpl,
      }),
      createStreamProvider({
        apiFormat: "openai",
        baseUrl: "  ",
        apiKey: key,
        model: "m",
        fetchImpl,
      }),
      createStreamProvider({
        apiFormat: "openai",
        baseUrl: "https://x",
        apiKey: key,
        model: "  ",
        fetchImpl,
      }),
      createStreamProvider({
        apiFormat: "openai",
        baseUrl: "https://x",
        apiKey: "   ",
        model: "m",
        fetchImpl,
      }),
    ];
    expect(cases[0]).toMatchObject({ ok: false, kind: "parse" });
    expect(cases[1]).toMatchObject({ ok: false, kind: "parse" });
    expect(cases[2]).toMatchObject({ ok: false, kind: "parse" });
    expect(cases[3]).toMatchObject({ ok: false, kind: "parse" });
    expect(cases[4]).toMatchObject({ ok: false, kind: "missing_key" });
    expect(fetchImpl).not.toHaveBeenCalled();
    for (const c of cases) {
      expect(JSON.stringify(c)).not.toContain(key);
    }
  });
});

describe("createStreamProviderFromStore", () => {
  it("no active provider fails without fetch", () => {
    const fetchImpl = vi.fn();
    const store = new ProviderStore();
    const result = createStreamProviderFromStore(store, { fetchImpl });
    expect(result).toMatchObject({ ok: false, kind: "parse" });
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("switching active openai → responses rebuilds provider (no singleton cache)", async () => {
    const store = new ProviderStore();
    const openai = store.add({
      name: "Chat",
      provider: "openai",
      apiKey: "sk-chat-key",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
      apiFormat: "openai",
    });
    const responses = store.add({
      name: "Resp",
      provider: "openai",
      apiKey: "sk-resp-key",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
      apiFormat: "responses",
    });
    store.setActive(openai.id);

    const urls: string[] = [];
    const fetchImpl = vi.fn(async (url: string) => {
      urls.push(url);
      if (url.includes("/responses")) return sseResponse(responsesDone);
      return sseResponse(openaiDone);
    });

    const first = createStreamProviderFromStore(store, { fetchImpl });
    expect(first.ok).toBe(true);
    if (!first.ok) return;
    await drain(first.provider);
    expect(urls[0]).toContain("/chat/completions");

    store.setActive(responses.id);
    const second = createStreamProviderFromStore(store, { fetchImpl });
    expect(second.ok).toBe(true);
    if (!second.ok) return;
    // Distinct instance (no shared singleton).
    expect(second.provider).not.toBe(first.provider);
    await drain(second.provider);
    expect(urls[1]).toContain("/responses");
    expect(fetchImpl).toHaveBeenCalledTimes(2);
  });

  it("active anthropic routes to /messages; empty key fails", async () => {
    const store = new ProviderStore();
    const ant = store.add({
      name: "Claude",
      provider: "anthropic",
      apiKey: "sk-ant-live",
      baseUrl: "https://api.anthropic.com/v1",
      model: "claude-test",
      apiFormat: "anthropic",
    });
    store.setActive(ant.id);
    const fetchImpl = vi.fn(async () => sseResponse(anthropicDone));
    const ok = createStreamProviderFromStore(store, { fetchImpl });
    expect(ok.ok).toBe(true);
    if (!ok.ok) return;
    await drain(ok.provider);
    const [url, init] = fetchImpl.mock.calls[0] as unknown as [
      string,
      RequestInit,
    ];
    expect(url).toBe("https://api.anthropic.com/v1/messages");
    expect((init.headers as Record<string, string>)["x-api-key"]).toBe(
      "sk-ant-live",
    );

    store.update(ant.id, { apiKey: "   " });
    const empty = createStreamProviderFromStore(store, { fetchImpl: vi.fn() });
    expect(empty).toMatchObject({ ok: false, kind: "missing_key" });
    expect(JSON.stringify(empty)).not.toContain("sk-ant-live");
  });
});

describe("createStreamProvider gateway mode stream", () => {
  for (const testCase of [
    {
      apiFormat: "openai",
      provider: "openai",
      upstreamId: "openai",
      suffix: "chat/completions",
      model: "gpt-4o",
      response: openaiDone,
      expectedReasoning: { reasoning: { effort: "xhigh" } },
    },
    {
      apiFormat: "responses",
      upstreamId: "responses",
      suffix: "responses",
      model: "gpt-4o",
      response: responsesDone,
      expectedReasoning: { reasoning: { effort: "xhigh", summary: "auto" } },
    },
    {
      apiFormat: "anthropic",
      upstreamId: "anthropic",
      suffix: "messages",
      model: "claude-test",
      response: anthropicDone,
      expectedReasoning: {
        thinking: { type: "adaptive" },
        output_config: { effort: "xhigh" },
      },
    },
  ] as const) {
    it(`${testCase.apiFormat} gateway streams without browser auth`, async () => {
      const browserSecret = "browser-secret-must-not-leak";
      const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
        expect(url).toBe(
          `https://app.example/api/ai/v1/${testCase.upstreamId}/${testCase.suffix}`,
        );
        const headers = init?.headers as Record<string, string>;
        expect(headers.Authorization).toBeUndefined();
        expect(headers["x-api-key"]).toBeUndefined();
        expect(headers.Accept).toBe("text/event-stream");
        expect(
          JSON.stringify({ headers, body: init?.body, url }),
        ).not.toContain(browserSecret);
        expect(JSON.parse(String(init?.body))).toMatchObject(
          testCase.expectedReasoning,
        );
        if (testCase.apiFormat === "anthropic") {
          expect(headers["anthropic-version"]).toBe("2023-06-01");
        }
        return sseResponse([...testCase.response]);
      });
      const created = createStreamProvider({
        apiFormat: testCase.apiFormat,
        provider:
          "provider" in testCase
            ? testCase.provider
            : testCase.apiFormat === "anthropic"
              ? "anthropic"
              : "openai",
        baseUrl: "https://vendor.example/v1",
        gatewayBaseUrl: "https://app.example",
        apiKey: browserSecret,
        model: testCase.model,
        connectionMode: "gateway",
        gatewayUpstreamId: testCase.upstreamId,
        reasoningMode: "max",
        fetchImpl,
      });
      expect(created.ok).toBe(true);
      if (!created.ok) return;
      await drain(created.provider);
      expect(fetchImpl).toHaveBeenCalledOnce();
    });
  }

  it("rejects gateway without upstream id", () => {
    const created = createStreamProvider({
      apiFormat: "openai",
      baseUrl: "https://api.openai.com/v1",
      gatewayBaseUrl: "https://app.example",
      apiKey: "",
      model: "gpt-4o",
      connectionMode: "gateway",
      gatewayUpstreamId: "",
    });
    expect(created.ok).toBe(false);
    if (!created.ok) expect(created.kind).toBe("parse");
  });

  it("rejects unsafe gateway URL, upstream id, and connection mode before fetch", () => {
    const fetchImpl = vi.fn();
    const inputs = [
      { gatewayBaseUrl: "//evil.example", gatewayUpstreamId: "openai" },
      {
        gatewayBaseUrl: "https://app.example/path",
        gatewayUpstreamId: "openai",
      },
      { gatewayBaseUrl: "", gatewayUpstreamId: "OpenAI" },
    ];
    for (const input of inputs) {
      expect(
        createStreamProvider({
          apiFormat: "openai",
          baseUrl: "https://api.openai.com/v1",
          apiKey: "",
          model: "gpt-4o",
          connectionMode: "gateway",
          ...input,
          fetchImpl,
        }),
      ).toMatchObject({ ok: false, kind: "parse" });
    }
    expect(
      createStreamProvider({
        apiFormat: "openai",
        baseUrl: "https://api.openai.com/v1",
        apiKey: "sk",
        model: "gpt-4o",
        connectionMode: "invalid",
        fetchImpl,
      }),
    ).toMatchObject({ ok: false, kind: "parse" });
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("createStreamProviderFromStore passes gateway fields", async () => {
    const store = new ProviderStore();
    const active = store.addFromTemplate("openai", "", {
      connectionMode: "gateway",
      gatewayUpstreamId: "openai",
      gatewayBaseUrl: "https://app.example",
    });
    store.update(active.id, { reasoningMode: "max" });
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      expect(url).toContain("/api/ai/v1/openai/chat/completions");
      expect(
        (init?.headers as Record<string, string>).Authorization,
      ).toBeUndefined();
      expect(JSON.parse(String(init?.body)).reasoning).toEqual({ effort: "xhigh" });
      return sseResponse(openaiDone);
    });
    const created = createStreamProviderFromStore(store, { fetchImpl });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    await drain(created.provider);
  });
});
