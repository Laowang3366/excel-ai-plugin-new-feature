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
