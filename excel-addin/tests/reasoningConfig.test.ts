import { describe, expect, it, vi } from "vitest";
import {
  applyAnthropicReasoningConfig,
  applyChatCompletionsReasoningConfig,
  applyResponsesReasoningConfig,
  isReasoningMode,
  resolveThinkingBudget,
} from "../shared/provider/reasoningConfig";
import { createStreamProvider } from "../shared/provider";
import type {
  AgentStreamEvent,
  StreamChatRequest,
} from "../shared/agent/types";
import type { ReasoningMode } from "../shared/provider/types";

const MODES: ReasoningMode[] = ["off", "low", "medium", "high", "max"];

const CHAT_VENDORS = [
  "openai",
  "deepseek",
  "zhipu",
  "kimi",
  "xiaomi",
  "aliyun",
  "tencent",
  "volcengine",
  "xunfei",
  "baidu",
  "jdcloud",
  "qwen",
  "minimax",
  "custom",
] as const;

function expectedChatBody(
  provider: string,
  mode: ReasoningMode,
): Record<string, unknown> {
  const body: Record<string, unknown> = {};
  applyChatCompletionsReasoningConfig(body, provider, mode);
  return body;
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

const openaiDone = [
  'data: {"id":"c","choices":[{"index":0,"delta":{"content":"ok"},"finish_reason":null}]}\n\n',
  'data: {"id":"c","choices":[{"index":0,"delta":{},"finish_reason":"stop"}]}\n\n',
  "data: [DONE]\n\n",
];

const responsesDone = [
  'data: {"type":"response.output_text.delta","item_id":"m1","delta":"ok"}\n\n',
  'data: {"type":"response.completed","response":{}}\n\n',
  "data: [DONE]\n\n",
];

const anthropicDone = [
  'data: {"type":"message_start","message":{"usage":{"input_tokens":1,"output_tokens":0}}}\n\n',
  'data: {"type":"content_block_start","index":0,"content_block":{"type":"text","text":"ok"}}\n\n',
  'data: {"type":"content_block_stop","index":0}\n\n',
  'data: {"type":"message_delta","delta":{"stop_reason":"end_turn"},"usage":{"output_tokens":1}}\n\n',
  'data: {"type":"message_stop"}\n\n',
  "data: [DONE]\n\n",
];

async function drainBody(
  apiFormat: "openai" | "responses" | "anthropic",
  options: {
    provider?: string;
    reasoningMode?: ReasoningMode | string;
    connectionMode?: "direct" | "gateway";
    apiKey?: string;
  },
): Promise<Record<string, unknown>> {
  let body: Record<string, unknown> = {};
  const secret = options.apiKey ?? "sk-must-not-leak";
  const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
    body = JSON.parse(String(init?.body));
    const chunks =
      apiFormat === "openai"
        ? openaiDone
        : apiFormat === "responses"
          ? responsesDone
          : anthropicDone;
    return sseResponse([...chunks]);
  });
  const created = createStreamProvider({
    apiFormat,
    provider: options.provider,
    baseUrl:
      apiFormat === "anthropic"
        ? "https://api.anthropic.com/v1"
        : "https://api.openai.com/v1",
    apiKey: secret,
    model: "test-model",
    connectionMode: options.connectionMode ?? "direct",
    gatewayBaseUrl:
      options.connectionMode === "gateway" ? "https://app.example" : undefined,
    gatewayUpstreamId:
      options.connectionMode === "gateway"
        ? apiFormat === "openai"
          ? "openai"
          : apiFormat
        : undefined,
    reasoningMode: options.reasoningMode as ReasoningMode | undefined,
    fetchImpl,
  });
  expect(created.ok).toBe(true);
  if (!created.ok) return body;
  const req: StreamChatRequest = {
    systemPrompt: "sys",
    messages: [{ role: "user", content: "hi" }],
    tools: [],
  };
  const events: AgentStreamEvent[] = [];
  for await (const e of created.provider.streamChat(req)) events.push(e);
  expect(JSON.stringify({ body, events })).not.toContain(secret);
  return body;
}

describe("reasoningConfig pure mapping", () => {
  it("validates reasoning modes and thinking budget", () => {
    expect(isReasoningMode("off")).toBe(true);
    expect(isReasoningMode("max")).toBe(true);
    expect(isReasoningMode("invalid")).toBe(false);
    expect(isReasoningMode(undefined)).toBe(false);
    expect(resolveThinkingBudget("low")).toBe(5000);
    expect(resolveThinkingBudget("medium")).toBe(10000);
    expect(resolveThinkingBudget("high")).toBe(20000);
    expect(resolveThinkingBudget("max")).toBe(20000);
    expect(resolveThinkingBudget(undefined)).toBe(20000);
  });

  it("responses: off → effort none; max → xhigh+summary; invalid omit", () => {
    const off: Record<string, unknown> = {};
    applyResponsesReasoningConfig(off, "off");
    expect(off).toEqual({ reasoning: { effort: "none" } });

    const max: Record<string, unknown> = {};
    applyResponsesReasoningConfig(max, "max");
    expect(max).toEqual({
      reasoning: { effort: "xhigh", summary: "auto" },
    });

    for (const mode of ["low", "medium", "high"] as const) {
      const body: Record<string, unknown> = {};
      applyResponsesReasoningConfig(body, mode);
      expect(body).toEqual({
        reasoning: { effort: mode, summary: "auto" },
      });
    }

    const invalid: Record<string, unknown> = {};
    applyResponsesReasoningConfig(invalid, "invalid");
    expect(invalid).toEqual({});
    applyResponsesReasoningConfig(invalid, undefined);
    expect(invalid).toEqual({});
  });

  it("anthropic: off/invalid omit; tiers map effort (max→xhigh)", () => {
    for (const mode of ["off", "invalid", undefined] as const) {
      const body: Record<string, unknown> = {};
      applyAnthropicReasoningConfig(body, mode as never);
      expect(body).toEqual({});
    }
    for (const mode of ["low", "medium", "high"] as const) {
      const body: Record<string, unknown> = {};
      applyAnthropicReasoningConfig(body, mode);
      expect(body).toEqual({
        thinking: { type: "adaptive" },
        output_config: { effort: mode },
      });
    }
    const max: Record<string, unknown> = {};
    applyAnthropicReasoningConfig(max, "max");
    expect(max).toEqual({
      thinking: { type: "adaptive" },
      output_config: { effort: "xhigh" },
    });
  });

  for (const provider of CHAT_VENDORS) {
    it(`chat vendor ${provider}: mode matrix + invalid defense`, () => {
      // off / invalid always empty
      for (const mode of ["off", "invalid", undefined] as const) {
        expect(expectedChatBody(provider, mode as never)).toEqual({});
      }

      const samples: ReasoningMode[] = ["low", "medium", "high", "max"];
      for (const mode of samples) {
        const body = expectedChatBody(provider, mode);
        switch (provider) {
          case "openai":
            expect(body).toEqual({
              reasoning: { effort: mode === "max" ? "xhigh" : mode },
            });
            expect(body).not.toHaveProperty("reasoning_effort");
            break;
          case "deepseek":
            expect(body).toEqual({
              thinking: { type: "enabled" },
              reasoning_effort: mode === "max" ? "max" : "high",
            });
            break;
          case "zhipu":
            expect(body).toEqual({
              thinking: { type: "enabled" },
              reasoning_effort:
                mode === "low" || mode === "medium"
                  ? "high"
                  : mode === "max"
                    ? "max"
                    : mode,
            });
            break;
          case "kimi":
            expect(body).toEqual({ thinking: { type: "enabled" } });
            break;
          case "xiaomi":
          case "tencent":
          case "jdcloud":
          case "custom":
            expect(body).toEqual({});
            break;
          case "xunfei":
          case "aliyun":
          case "qwen":
            expect(body).toEqual({
              enable_thinking: true,
              thinking_budget: resolveThinkingBudget(mode),
            });
            break;
          case "baidu":
            expect(body).toEqual({
              enable_search: false,
              enable_citation: false,
            });
            break;
          case "volcengine":
            expect(body).toEqual({
              thinking: {
                type: "enabled",
                budget_tokens: resolveThinkingBudget(mode),
              },
            });
            break;
          case "minimax":
            expect(body).toEqual({
              reasoning_config: {
                effort: mode === "max" ? "xhigh" : mode,
              },
            });
            break;
          default:
            expect(body).toEqual({});
        }
      }
    });
  }

  it("unknown provider never receives OpenAI-only fields", () => {
    for (const mode of MODES) {
      if (mode === "off") continue;
      expect(expectedChatBody("not-a-vendor", mode)).toEqual({});
      expect(expectedChatBody("", mode)).toEqual({});
    }
  });
});

describe("reasoningConfig via createStreamProvider (direct + gateway)", () => {
  it.each([
    ["deepseek", "high", { thinking: { type: "enabled" }, reasoning_effort: "high" }],
    ["deepseek", "max", { thinking: { type: "enabled" }, reasoning_effort: "max" }],
    ["openai", "max", { reasoning: { effort: "xhigh" } }],
    ["kimi", "high", { thinking: { type: "enabled" } }],
    ["xiaomi", "high", {}],
    ["custom", "high", {}],
    ["minimax", "max", { reasoning_config: { effort: "xhigh" } }],
    ["qwen", "low", { enable_thinking: true, thinking_budget: 5000 }],
    ["volcengine", "medium", { thinking: { type: "enabled", budget_tokens: 10000 } }],
    ["baidu", "high", { enable_search: false, enable_citation: false }],
  ] as const)(
    "direct openai provider=%s mode=%s",
    async (provider, mode, expectedPatch) => {
      const body = await drainBody("openai", {
        provider,
        reasoningMode: mode,
        connectionMode: "direct",
      });
      for (const [k, v] of Object.entries(expectedPatch)) {
        expect(body[k]).toEqual(v);
      }
      if (Object.keys(expectedPatch).length === 0) {
        expect(body).not.toHaveProperty("reasoning");
        expect(body).not.toHaveProperty("reasoning_effort");
        expect(body).not.toHaveProperty("thinking");
        expect(body).not.toHaveProperty("enable_thinking");
      }
      // secret already checked in drainBody
    },
  );

  it("gateway reuses the same chat body mapping without browser auth headers", async () => {
    let capturedUrl = "";
    let capturedHeaders: Record<string, string> = {};
    let body: Record<string, unknown> = {};
    const browserSecret = "browser-secret-must-not-leak";
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      capturedUrl = url;
      capturedHeaders = (init?.headers ?? {}) as Record<string, string>;
      body = JSON.parse(String(init?.body));
      return sseResponse([...openaiDone]);
    });
    const created = createStreamProvider({
      apiFormat: "openai",
      provider: "deepseek",
      baseUrl: "https://vendor.example/v1",
      gatewayBaseUrl: "https://app.example",
      apiKey: browserSecret,
      model: "deepseek-chat",
      connectionMode: "gateway",
      gatewayUpstreamId: "deepseek",
      reasoningMode: "max",
      fetchImpl,
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    for await (const _ of created.provider.streamChat({
      systemPrompt: "",
      messages: [{ role: "user", content: "hi" }],
      tools: [],
    })) {
      /* drain */
    }
    expect(capturedUrl).toBe(
      "https://app.example/api/ai/v1/deepseek/chat/completions",
    );
    expect(capturedHeaders.Authorization).toBeUndefined();
    expect(capturedHeaders["x-api-key"]).toBeUndefined();
    expect(body.thinking).toEqual({ type: "enabled" });
    expect(body.reasoning_effort).toBe("max");
    expect(JSON.stringify({ body, headers: capturedHeaders })).not.toContain(
      browserSecret,
    );
  });

  it("responses gateway off sends effort none; max sends xhigh", async () => {
    const off = await drainBody("responses", {
      provider: "openai",
      reasoningMode: "off",
      connectionMode: "gateway",
      apiKey: "gateway-browser-secret",
    });
    expect(off.reasoning).toEqual({ effort: "none" });

    const max = await drainBody("responses", {
      provider: "openai",
      reasoningMode: "max",
      connectionMode: "gateway",
      apiKey: "gateway-browser-secret",
    });
    expect(max.reasoning).toEqual({ effort: "xhigh", summary: "auto" });
  });

  it("anthropic gateway max maps xhigh; off omits", async () => {
    const max = await drainBody("anthropic", {
      provider: "anthropic",
      reasoningMode: "max",
      connectionMode: "gateway",
    });
    expect(max.thinking).toEqual({ type: "adaptive" });
    expect(max.output_config).toEqual({ effort: "xhigh" });

    const off = await drainBody("anthropic", {
      provider: "anthropic",
      reasoningMode: "off",
      connectionMode: "gateway",
    });
    expect(off).not.toHaveProperty("thinking");
    expect(off).not.toHaveProperty("output_config");
  });
});
