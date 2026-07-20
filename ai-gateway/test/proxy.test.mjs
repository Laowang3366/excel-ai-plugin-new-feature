import assert from "node:assert/strict";
import { describe, it } from "node:test";
import {
  filterRequestHeaders,
  filterResponseHeaders,
  joinUpstreamUrl,
} from "../src/proxy.mjs";

describe("proxy header and path guards", () => {
  it("drops client credentials and hop-by-hop headers", () => {
    const out = filterRequestHeaders({
      authorization: "Bearer client-secret",
      "x-api-key": "client-secret",
      cookie: "a=1",
      host: "evil.example",
      connection: "keep-alive",
      "content-type": "application/json",
      accept: "text/event-stream",
      "anthropic-version": "2023-06-01",
      "openai-organization": "org",
      "x-custom": "nope",
    });
    assert.deepEqual(out, {
      "content-type": "application/json",
      accept: "text/event-stream",
      "anthropic-version": "2023-06-01",
      "openai-organization": "org",
    });
  });

  it("strips CR/LF from forwarded values", () => {
    const out = filterRequestHeaders({
      "content-type": "application/json\r\nX-Injected: 1",
      accept: "application/json",
    });
    assert.equal(out["content-type"], undefined);
    assert.equal(out.accept, "application/json");
  });

  it("only allowlists safe response headers", () => {
    const headers = new Headers({
      "content-type": "text/event-stream",
      "set-cookie": "secret=1",
      "x-request-id": "r1",
      server: "upstream",
    });
    assert.deepEqual(filterResponseHeaders(headers), {
      "content-type": "text/event-stream",
      "x-request-id": "r1",
    });
  });

  it("joins only fixed endpoint suffixes", () => {
    assert.equal(joinUpstreamUrl("https://api.example/v1", "chat/completions"), "https://api.example/v1/chat/completions");
    assert.equal(joinUpstreamUrl("https://api.example/v1", "responses"), "https://api.example/v1/responses");
    assert.equal(joinUpstreamUrl("https://api.example/v1", "messages"), "https://api.example/v1/messages");
    assert.equal(joinUpstreamUrl("https://api.example/v1", "models"), "https://api.example/v1/models");
    assert.throws(() => joinUpstreamUrl("https://api.example/v1", "chat/completions/../evil"), /unsupported/);
    assert.throws(() => joinUpstreamUrl("https://api.example/v1", ""), /unsupported/);
  });
});
