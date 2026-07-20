import { describe, expect, it } from "vitest";
import {
  assertSafeBaseUrl,
  assertSafeGatewayUpstreamId,
  resolveProviderEndpoint,
} from "../shared/provider";

describe("endpointResolve", () => {
  it("resolves direct openai/responses/anthropic endpoints with auth", () => {
    const chat = resolveProviderEndpoint({
      connectionMode: "direct",
      baseUrl: "https://api.openai.com/v1",
      apiKey: "sk-direct",
      apiFormat: "openai",
      kind: "chat",
      acceptEventStream: true,
    });
    expect(chat.ok).toBe(true);
    if (chat.ok) {
      expect(chat.data.url).toBe("https://api.openai.com/v1/chat/completions");
      expect(chat.data.headers.Authorization).toBe("Bearer sk-direct");
      expect(chat.data.headers.Accept).toBe("text/event-stream");
    }

    const responses = resolveProviderEndpoint({
      connectionMode: "direct",
      baseUrl: "https://api.openai.com/v1/",
      apiKey: "sk-r",
      apiFormat: "responses",
      kind: "responses",
    });
    expect(responses.ok).toBe(true);
    if (responses.ok) {
      expect(responses.data.url).toBe("https://api.openai.com/v1/responses");
      expect(responses.data.headers.Authorization).toBe("Bearer sk-r");
    }

    const anthropic = resolveProviderEndpoint({
      connectionMode: "direct",
      baseUrl: "https://api.anthropic.com/v1",
      apiKey: "ak",
      apiFormat: "anthropic",
      kind: "messages",
    });
    expect(anthropic.ok).toBe(true);
    if (anthropic.ok) {
      expect(anthropic.data.url).toBe("https://api.anthropic.com/v1/messages");
      expect(anthropic.data.headers["x-api-key"]).toBe("ak");
      expect(anthropic.data.headers["anthropic-version"]).toBe("2023-06-01");
    }
  });

  it("requires api key in direct mode", () => {
    const r = resolveProviderEndpoint({
      connectionMode: "direct",
      baseUrl: "https://api.openai.com/v1",
      apiKey: "",
      apiFormat: "openai",
      kind: "chat",
    });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.kind).toBe("missing_key");
  });

  it("builds gateway paths for all formats without auth headers", () => {
    for (const [format, kind, suffix] of [
      ["openai", "chat", "chat/completions"],
      ["responses", "responses", "responses"],
      ["anthropic", "messages", "messages"],
      ["openai", "models", "models"],
    ] as const) {
      const r = resolveProviderEndpoint({
        connectionMode: "gateway",
        baseUrl: "https://vendor.example/v1",
        gatewayBaseUrl: "https://plugin.example.com",
        apiKey: "should-not-appear",
        apiFormat: format,
        gatewayUpstreamId: "openai",
        kind,
        acceptEventStream: kind !== "models",
      });
      expect(r.ok).toBe(true);
      if (!r.ok) continue;
      expect(r.data.url).toBe(
        `https://plugin.example.com/api/ai/v1/openai/${suffix}`,
      );
      expect(r.data.headers.Authorization).toBeUndefined();
      expect(r.data.headers["x-api-key"]).toBeUndefined();
      expect(r.data.headers.authorization).toBeUndefined();
      expect(JSON.stringify(r.data)).not.toContain("should-not-appear");
      if (format === "anthropic") {
        expect(r.data.headers["anthropic-version"]).toBe("2023-06-01");
      }
    }
  });

  it("allows empty gatewayBaseUrl for same-origin", () => {
    const r = resolveProviderEndpoint({
      connectionMode: "gateway",
      baseUrl: "https://vendor.example/v1",
      gatewayBaseUrl: "",
      apiFormat: "openai",
      gatewayUpstreamId: "fake",
      kind: "models",
    });
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.data.url).toBe("/api/ai/v1/fake/models");
  });

  it("rejects unsafe upstream ids and base urls", () => {
    expect(assertSafeGatewayUpstreamId("").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("a/b").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("a?x").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("a#h").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("../x").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("OpenAI").ok).toBe(false);
    expect(assertSafeGatewayUpstreamId("openai").ok).toBe(true);

    expect(
      assertSafeBaseUrl("https://user:pass@evil/v1", "direct").ok,
    ).toBe(false);
    expect(assertSafeBaseUrl("https://api.openai.com/v1?x=1", "direct").ok).toBe(
      false,
    );
    expect(assertSafeBaseUrl("https://api.openai.com/v1#h", "direct").ok).toBe(
      false,
    );

    const bad = resolveProviderEndpoint({
      connectionMode: "gateway",
      baseUrl: "https://vendor.example/v1",
      gatewayBaseUrl: "https://plugin.example.com",
      apiFormat: "openai",
      gatewayUpstreamId: "openai/../admin",
      kind: "chat",
    });
    expect(bad.ok).toBe(false);

    for (const gatewayBaseUrl of [
      "//evil.example",
      "https://plugin.example.com/path",
      "https://user:pass@plugin.example.com",
      "https://plugin.example.com?x=1",
      "https://plugin.example.com#fragment",
    ]) {
      const result = resolveProviderEndpoint({
        connectionMode: "gateway",
        baseUrl: "https://vendor.example/v1",
        gatewayBaseUrl,
        apiFormat: "openai",
        gatewayUpstreamId: "openai",
        kind: "chat",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.kind).toBe("parse");
    }

    const unknownMode = resolveProviderEndpoint({
      connectionMode: "proxy",
      baseUrl: "https://vendor.example/v1",
      apiFormat: "openai",
      gatewayUpstreamId: "openai",
      kind: "chat",
    });
    expect(unknownMode.ok).toBe(false);
    if (!unknownMode.ok) expect(unknownMode.kind).toBe("parse");
  });
});
