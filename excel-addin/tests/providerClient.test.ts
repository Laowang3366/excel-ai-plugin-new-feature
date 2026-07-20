import { describe, expect, it, vi } from "vitest";
import {
  ProviderClient,
  buildListModelsRequest,
  buildTestConnectionRequest,
} from "../shared/provider";

describe("ProviderClient request contracts", () => {
  it("builds openai chat completions test body", () => {
    const req = buildTestConnectionRequest({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-test",
      apiFormat: "openai",
      model: "gpt-5.4",
    });
    expect(req.url).toBe("https://api.openai.com/v1/chat/completions");
    expect(req.headers.Authorization).toBe("Bearer sk-test");
    expect(req.body).toMatchObject({
      model: "gpt-5.4",
      max_tokens: 1,
      messages: [{ role: "user", content: "Hi" }],
    });
  });

  it("builds responses and anthropic contracts", () => {
    const responses = buildTestConnectionRequest({
      baseUrl: "https://api.openai.com/v1/",
      apiKey: "sk-r",
      apiFormat: "responses",
      model: "gpt-5.4",
    });
    expect(responses.url).toBe("https://api.openai.com/v1/responses");
    expect(responses.body).toMatchObject({ input: "Hi", max_output_tokens: 1 });

    const anthropic = buildTestConnectionRequest({
      baseUrl: "https://api.anthropic.com/v1",
      apiKey: "ak",
      apiFormat: "anthropic",
      model: "claude-sonnet-4-6",
    });
    expect(anthropic.url).toBe("https://api.anthropic.com/v1/messages");
    expect(anthropic.headers["x-api-key"]).toBe("ak");
    expect(anthropic.headers["anthropic-version"]).toBe("2023-06-01");
    expect(anthropic.body).toMatchObject({ max_tokens: 1 });
  });

  it("listModels uses /models and rejects anthropic", () => {
    const openai = buildListModelsRequest({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk",
      apiFormat: "openai",
    });
    expect(openai?.url).toBe("https://api.openai.com/v1/models");
    expect(
      buildListModelsRequest({
        baseUrl: "https://api.anthropic.com/v1",
        apiKey: "ak",
        apiFormat: "anthropic",
      }),
    ).toBeNull();
  });
});

describe("ProviderClient fetch behavior", () => {
  it("fails explicitly without api key", async () => {
    const client = new ProviderClient(vi.fn());
    const result = await client.testConnection({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "",
      apiFormat: "openai",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.kind).toBe("missing_key");
  });

  it("keeps the direct missing-base error contract", async () => {
    const fetchImpl = vi.fn();
    const client = new ProviderClient(fetchImpl);
    for (const result of [
      await client.testConnection({
        baseUrl: "",
        apiKey: "sk",
        apiFormat: "openai",
      }),
      await client.listModels({
        baseUrl: "",
        apiKey: "ak",
        apiFormat: "anthropic",
      }),
    ]) {
      expect(result).toMatchObject({
        ok: false,
        kind: "http",
        error: "Base URL 未设置",
      });
    }
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it("returns http error details", async () => {
    const fetchImpl = vi.fn(async () =>
      new Response(JSON.stringify({ error: { message: "bad key" } }), {
        status: 401,
        headers: { "Content-Type": "application/json" },
      }),
    );
    const client = new ProviderClient(fetchImpl);
    const result = await client.testConnection({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk",
      apiFormat: "openai",
      model: "gpt-5.4",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.kind).toBe("http");
      expect(result.error).toContain("bad key");
      expect(result.status).toBe(401);
    }
    expect(JSON.stringify(result)).not.toContain("sk");
  });

  it("redacts the normalized direct key from upstream errors", async () => {
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).Authorization).toBe(
        "Bearer sk-trimmed-secret",
      );
      return new Response(
        JSON.stringify({ error: { message: "rejected sk-trimmed-secret" } }),
        { status: 401, headers: { "Content-Type": "application/json" } },
      );
    });
    const client = new ProviderClient(fetchImpl);
    const result = await client.testConnection({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "  sk-trimmed-secret  ",
      apiFormat: "openai",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toContain("[REDACTED]");
    expect(JSON.stringify(result)).not.toContain("sk-trimmed-secret");
  });

  it("classifies failed fetch as cors/network", async () => {
    const fetchImpl = vi.fn(async () => {
      throw new TypeError("Failed to fetch");
    });
    const client = new ProviderClient(fetchImpl);
    const result = await client.listModels({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk",
      apiFormat: "openai",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.kind).toBe("cors");
      expect(result.error.toLowerCase()).toContain("cors");
    }
  });

  it("parses openai models list", async () => {
    const fetchImpl = vi.fn(async () =>
      new Response(JSON.stringify({ data: [{ id: "b" }, { id: "a" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      }),
    );
    const client = new ProviderClient(fetchImpl);
    const result = await client.listModels({
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk",
      apiFormat: "openai",
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.data.models).toEqual(["a", "b"]);
  });

  it("marks anthropic listModels unsupported", async () => {
    const client = new ProviderClient(vi.fn());
    const result = await client.listModels({
      baseUrl: "https://api.anthropic.com/v1",
      apiKey: "ak",
      apiFormat: "anthropic",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.kind).toBe("unsupported");
  });
});


describe("ProviderClient gateway mode", () => {
  it("testConnection posts gateway path without auth headers", async () => {
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      expect(url).toBe("https://app.example/api/ai/v1/openai/chat/completions");
      const headers = init?.headers as Record<string, string>;
      expect(headers.Authorization).toBeUndefined();
      expect(headers["x-api-key"]).toBeUndefined();
      expect(headers["Content-Type"]).toBe("application/json");
      expect(String(init?.body)).not.toContain("sk-");
      return new Response("{}", { status: 200 });
    });
    const client = new ProviderClient(fetchImpl);
    const result = await client.testConnection({
      baseUrl: "https://api.openai.com/v1",
      gatewayBaseUrl: "https://app.example",
      apiKey: "",
      apiFormat: "openai",
      model: "gpt-4o",
      connectionMode: "gateway",
      gatewayUpstreamId: "openai",
    });
    expect(result.ok).toBe(true);
    expect(fetchImpl).toHaveBeenCalledOnce();
  });

  it("listModels uses gateway models path without auth", async () => {
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      expect(url).toBe("/api/ai/v1/openai/models");
      const headers = init?.headers as Record<string, string>;
      expect(headers.Authorization).toBeUndefined();
      expect(headers["x-api-key"]).toBeUndefined();
      return new Response(JSON.stringify({ data: [{ id: "m1" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    });
    const client = new ProviderClient(fetchImpl);
    const result = await client.listModels({
      baseUrl: "https://api.openai.com/v1",
      gatewayBaseUrl: "",
      apiKey: "leak-me",
      apiFormat: "openai",
      connectionMode: "gateway",
      gatewayUpstreamId: "openai",
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.data.models).toEqual(["m1"]);
    expect(JSON.stringify(fetchImpl.mock.calls[0]?.[1]?.headers)).not.toContain(
      "leak-me",
    );
  });

  it("allows Anthropic model listing through the gateway", async () => {
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      expect(url).toBe("https://app.example/api/ai/v1/anthropic/models");
      const headers = init?.headers as Record<string, string>;
      expect(headers.Authorization).toBeUndefined();
      expect(headers["x-api-key"]).toBeUndefined();
      expect(headers["anthropic-version"]).toBe("2023-06-01");
      return new Response(JSON.stringify({ data: [{ id: "claude-test" }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    });
    const client = new ProviderClient(fetchImpl);
    const result = await client.listModels({
      baseUrl: "https://api.anthropic.com/v1",
      gatewayBaseUrl: "https://app.example",
      apiKey: "",
      apiFormat: "anthropic",
      connectionMode: "gateway",
      gatewayUpstreamId: "anthropic",
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.data.models).toEqual(["claude-test"]);
  });

  it("rejects illegal gatewayUpstreamId fail-closed", async () => {
    const client = new ProviderClient(vi.fn());
    const result = await client.testConnection({
      baseUrl: "https://api.openai.com/v1",
      gatewayBaseUrl: "https://app.example",
      apiKey: "",
      apiFormat: "openai",
      connectionMode: "gateway",
      gatewayUpstreamId: "openai/../x",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.kind).toBe("parse");
  });
});
