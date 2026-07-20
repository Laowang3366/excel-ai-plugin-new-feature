import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { request, startFakeUpstream, startGateway } from "./helpers.mjs";

describe("ai-gateway http", () => {
  it("GET /healthz", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const res = await request(`${gw.base}/healthz`);
      assert.equal(res.status, 200);
      assert.deepEqual(res.json(), { status: "ok" });
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("GET models proxies and injects auth", async () => {
    const up = await startFakeUpstream({
      body: JSON.stringify({ data: [{ id: "m1" }] }),
    });
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/models`, {
        headers: {
          authorization: "Bearer client-should-not-win",
          cookie: "a=1",
          host: "evil.example",
          accept: "application/json",
          "x-custom": "nope",
        },
      });
      assert.equal(res.status, 200);
      assert.deepEqual(res.json(), { data: [{ id: "m1" }] });
      assert.equal(up.requests.length, 1);
      const hit = up.requests[0];
      assert.equal(hit.method, "GET");
      assert.equal(hit.url, "/v1/models");
      assert.equal(hit.headers.authorization, "Bearer test-secret-key");
      assert.equal(hit.headers.cookie, undefined);
      assert.equal(hit.headers["x-custom"], undefined);
      // response header allowlist
      assert.equal(res.headers["x-request-id"], "fake-req-1");
      assert.equal(res.headers["set-cookie"], undefined);
      assert.equal(res.headers["x-should-strip"], undefined);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("POST chat/completions, responses, messages", async () => {
    const up = await startFakeUpstream({ body: JSON.stringify({ id: "c1" }) });
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      for (const suffix of ["chat/completions", "responses", "messages"]) {
        const res = await request(`${gw.base}/api/ai/v1/fake/${suffix}`, {
          method: "POST",
          headers: {
            "content-type": "application/json",
            "anthropic-version": "2023-06-01",
            "openai-organization": "org-x",
            authorization: "Bearer evil",
          },
          body: JSON.stringify({ model: "x", stream: false }),
        });
        assert.equal(res.status, 200, suffix);
        assert.equal(res.json().id, "c1");
      }
      assert.equal(up.requests.length, 3);
      for (const hit of up.requests) {
        assert.equal(hit.headers.authorization, "Bearer test-secret-key");
        assert.equal(hit.headers["anthropic-version"], "2023-06-01");
        assert.equal(hit.headers["openai-organization"], "org-x");
        assert.match(hit.body.toString("utf8"), /"model":"x"/);
      }
      assert.equal(up.requests[0].url, "/v1/chat/completions");
      assert.equal(up.requests[1].url, "/v1/responses");
      assert.equal(up.requests[2].url, "/v1/messages");
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("streams bytes transparently", async () => {
    const chunks = ["data: {\"a\":1}\n\n", "data: [DONE]\n\n"];
    const up = await startFakeUpstream({
      headers: { "content-type": "text/event-stream" },
      streamChunks: chunks,
    });
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const res = await fetch(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ stream: true }),
      });
      assert.equal(res.status, 200);
      assert.match(res.headers.get("content-type") || "", /text\/event-stream/);
      const text = await res.text();
      assert.equal(text, chunks.join(""));
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("rejects body over 4 MiB with 413 before proxy", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const big = Buffer.alloc(4 * 1024 * 1024 + 1, 0x61);
      // wrap as JSON string roughly - actually need valid json; use raw oversized
      // server checks size before full parse path; content-type json with huge body
      const res = await fetch(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: big,
      });
      assert.equal(res.status, 413);
      assert.equal(up.requests.length, 0);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("rejects unknown upstream and unknown path", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const u = await request(`${gw.base}/api/ai/v1/nope/models`);
      assert.equal(u.status, 404);
      assert.equal(u.json().error, "unknown_upstream");

      const p = await request(`${gw.base}/api/ai/v1/fake/not-a-path`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: "{}",
      });
      assert.equal(p.status, 404);

      const root = await request(`${gw.base}/api/other`);
      assert.equal(root.status, 404);
      assert.equal(up.requests.length, 0);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("rejects non-json content type and invalid json", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const ct = await request(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "text/plain" },
        body: "{}",
      });
      assert.equal(ct.status, 415);

      const bad = await request(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: "not-json",
      });
      assert.equal(bad.status, 400);
      assert.equal(up.requests.length, 0);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("rate limits fail-closed", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        AI_GATEWAY_RATE_LIMIT_MAX: "2",
        AI_GATEWAY_RATE_LIMIT_WINDOW_MS: "60000",
      },
    });
    try {
      assert.equal((await request(`${gw.base}/healthz`)).status, 200);
      // healthz also counts? Currently rate limit only on API paths.
      // Force API:
      assert.equal((await request(`${gw.base}/api/ai/v1/fake/models`)).status, 200);
      assert.equal((await request(`${gw.base}/api/ai/v1/fake/models`)).status, 200);
      const limited = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(limited.status, 429);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("concurrency limits fail-closed", async () => {
    const up = await startFakeUpstream({ delayMs: 200 });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        AI_GATEWAY_MAX_CONCURRENT: "1",
        AI_GATEWAY_RATE_LIMIT_MAX: "1000",
      },
    });
    try {
      const p1 = request(`${gw.base}/api/ai/v1/fake/models`);
      await new Promise((r) => setTimeout(r, 20));
      const p2 = request(`${gw.base}/api/ai/v1/fake/models`);
      const [r1, r2] = await Promise.all([p1, p2]);
      const statuses = [r1.status, r2.status].sort();
      assert.deepEqual(statuses, [200, 503]);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("times out hung upstream", async () => {
    const up = await startFakeUpstream({ hang: true });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        AI_GATEWAY_CONNECT_TIMEOUT_MS: "100",
        AI_GATEWAY_TOTAL_TIMEOUT_MS: "200",
      },
    });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(res.status, 504);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("aborts upstream when client disconnects mid-stream", async () => {
    let upstreamFetchAborted = false;
    const up = await startFakeUpstream({ hang: true });

    const fetchImpl = async (_url, init) => {
      await new Promise((resolve, reject) => {
        const signal = init?.signal;
        if (!signal) {
          reject(new Error("missing signal"));
          return;
        }
        if (signal.aborted) {
          upstreamFetchAborted = true;
          reject(Object.assign(new Error("aborted"), { name: "AbortError" }));
          return;
        }
        signal.addEventListener(
          "abort",
          () => {
            upstreamFetchAborted = true;
            reject(Object.assign(new Error("aborted"), { name: "AbortError" }));
          },
          { once: true },
        );
      });
      throw new Error("unreachable");
    };

    const gw = await startGateway({
      upstreamPort: up.port,
      fetchImpl,
      env: {
        AI_GATEWAY_CONNECT_TIMEOUT_MS: "5000",
        AI_GATEWAY_TOTAL_TIMEOUT_MS: "8000",
      },
    });
    try {
      const http = await import("node:http");
      await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("disconnect timeout")), 3000);
        const req = http.request(
          {
            hostname: "127.0.0.1",
            port: gw.port,
            path: "/api/ai/v1/fake/models",
            method: "GET",
          },
          () => {
            // should not get response before abort
          },
        );
        req.on("error", () => {
          clearTimeout(timer);
          resolve();
        });
        req.end();
        // Disconnect client while gateway is waiting on hung upstream.
        setTimeout(() => {
          req.destroy();
        }, 50);
      });
      const deadline = Date.now() + 2000;
      while (!upstreamFetchAborted && Date.now() < deadline) {
        await new Promise((r) => setTimeout(r, 20));
      }
      assert.equal(upstreamFetchAborted, true);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("uses x-api-key auth type", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        ANTHROPIC_API_KEY: "ant-secret",
        AI_GATEWAY_UPSTREAMS_JSON: JSON.stringify({
          fake: {
            baseUrl: `http://127.0.0.1:${up.port}/v1`,
            auth: { type: "x-api-key", env: "ANTHROPIC_API_KEY" },
          },
        }),
      },
    });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(res.status, 200);
      assert.equal(up.requests[0].headers["x-api-key"], "ant-secret");
      assert.equal(up.requests[0].headers.authorization, undefined);
    } finally {
      await gw.close();
      await up.close();
    }
  });
});
