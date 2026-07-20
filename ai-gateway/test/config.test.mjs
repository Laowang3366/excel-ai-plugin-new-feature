import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { assertSafeBaseUrl, loadConfig } from "../src/config.mjs";

function baseEnv(overrides = {}) {
  return {
    AI_GATEWAY_ALLOW_LOCAL_UPSTREAMS: "0",
    OPENAI_API_KEY: "sk-test",
    AI_GATEWAY_UPSTREAMS_JSON: JSON.stringify({
      openai: {
        baseUrl: "https://api.openai.com/v1",
        auth: { type: "bearer", env: "OPENAI_API_KEY" },
      },
    }),
    ...overrides,
  };
}

describe("config", () => {
  it("loads valid https upstreams", () => {
    const cfg = loadConfig(baseEnv());
    assert.equal(cfg.upstreams.size, 1);
    assert.equal(cfg.upstreams.get("openai").baseUrl, "https://api.openai.com/v1");
    assert.equal(cfg.maxBodyBytes, 4 * 1024 * 1024);
  });

  it("rejects missing UPSTREAMS_JSON", () => {
    assert.throws(() => loadConfig({}), /AI_GATEWAY_UPSTREAMS_JSON is required/);
  });

  it("rejects invalid JSON", () => {
    assert.throws(
      () => loadConfig({ AI_GATEWAY_UPSTREAMS_JSON: "{", OPENAI_API_KEY: "x" }),
      /valid JSON/,
    );
  });

  it("rejects empty upstream map", () => {
    assert.throws(
      () => loadConfig({ AI_GATEWAY_UPSTREAMS_JSON: "{}", OPENAI_API_KEY: "x" }),
      /at least one/,
    );
  });

  it("rejects missing secret env", () => {
    assert.throws(() => loadConfig(baseEnv({ OPENAI_API_KEY: "" })), /missing or empty/);
  });

  it("rejects credentials in baseUrl", () => {
    assert.throws(
      () =>
        assertSafeBaseUrl("https://user:pass@api.openai.com/v1", {
          allowLocalUpstreams: false,
          id: "x",
        }),
      /credentials/,
    );
  });

  it("rejects fragment in baseUrl", () => {
    assert.throws(
      () =>
        assertSafeBaseUrl("https://api.openai.com/v1#frag", {
          allowLocalUpstreams: false,
          id: "x",
        }),
      /fragment/,
    );
  });

  it("rejects private https host by default", () => {
    assert.throws(
      () =>
        assertSafeBaseUrl("https://127.0.0.1/v1", {
          allowLocalUpstreams: false,
          id: "x",
        }),
      /private|local/,
    );
  });

  it("rejects http without local flag", () => {
    assert.throws(
      () =>
        assertSafeBaseUrl("http://127.0.0.1:9/v1", {
          allowLocalUpstreams: false,
          id: "x",
        }),
      /ALLOW_LOCAL/,
    );
  });

  it("allows loopback http when local flag set", () => {
    const url = assertSafeBaseUrl("http://127.0.0.1:9/v1/", {
      allowLocalUpstreams: true,
      id: "x",
    });
    assert.equal(url, "http://127.0.0.1:9/v1");
  });

  it("rejects non-loopback http even with local flag", () => {
    assert.throws(
      () =>
        assertSafeBaseUrl("http://example.com/v1", {
          allowLocalUpstreams: true,
          id: "x",
        }),
      /loopback/,
    );
  });

  it("rejects invalid upstream id", () => {
    assert.throws(
      () =>
        loadConfig(
          baseEnv({
            AI_GATEWAY_UPSTREAMS_JSON: JSON.stringify({
              "Bad Id": {
                baseUrl: "https://api.openai.com/v1",
                auth: { type: "bearer", env: "OPENAI_API_KEY" },
              },
            }),
          }),
        ),
      /invalid upstream id/,
    );
  });
});
