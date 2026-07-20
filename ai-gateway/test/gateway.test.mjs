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


  it("total timeout covers slow response body after headers", async () => {
    const up = await startFakeUpstream({ hangAfterHeaders: true, hangAfterHeadersWrite: "partial" });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        AI_GATEWAY_CONNECT_TIMEOUT_MS: "100",
        AI_GATEWAY_TOTAL_TIMEOUT_MS: "150",
      },
    });
    try {
      const started = Date.now();
      let sawHeaders = false;
      let closed = false;
      const http = await import("node:http");
      await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("body total timeout wait")), 3000);
        const req = http.request(
          {
            hostname: "127.0.0.1",
            port: gw.port,
            path: "/api/ai/v1/fake/models",
            method: "GET",
          },
          (res) => {
            sawHeaders = true;
            // headers should arrive (200 from upstream) before total timeout kills stream
            assert.equal(res.statusCode, 200);
            res.on("data", () => {});
            res.on("close", () => {
              closed = true;
              clearTimeout(timer);
              resolve();
            });
            res.on("end", () => {
              closed = true;
              clearTimeout(timer);
              resolve();
            });
          },
        );
        req.on("error", () => {
          closed = true;
          clearTimeout(timer);
          resolve();
        });
        req.end();
      });
      assert.equal(sawHeaders, true);
      assert.equal(closed, true);
      assert.ok(Date.now() - started < 2500);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("drain wait wakes on client abort and releases concurrency", async () => {
    // Upstream sends one huge chunk so gateway hits backpressure (needs drain).
    const huge = "x".repeat(1024 * 256);
    const up = await startFakeUpstream({
      headers: { "content-type": "application/octet-stream" },
      streamChunks: [huge, huge, huge, huge],
    });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        AI_GATEWAY_MAX_CONCURRENT: "1",
        AI_GATEWAY_RATE_LIMIT_MAX: "1000",
        AI_GATEWAY_CONNECT_TIMEOUT_MS: "2000",
        AI_GATEWAY_TOTAL_TIMEOUT_MS: "5000",
      },
    });
    try {
      const http = await import("node:http");
      // Start a slow consumer that will abort mid-stream
      const first = new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("first request hung")), 3000);
        const req = http.request(
          {
            hostname: "127.0.0.1",
            port: gw.port,
            path: "/api/ai/v1/fake/chat/completions",
            method: "POST",
            headers: {
              "content-type": "application/json",
              "content-length": 2,
            },
          },
          (res) => {
            // pause reading to encourage buffer fill, then abort
            setTimeout(() => {
              req.destroy();
              clearTimeout(timer);
              resolve(res.statusCode);
            }, 30);
          },
        );
        req.on("error", () => {
          clearTimeout(timer);
          resolve("aborted");
        });
        req.write("{}");
        req.end();
      });
      await first;
      // Concurrency slot must be free for a subsequent request.
      const second = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(second.status, 200);
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

  it("forwards upstream non-2xx status and body without remapping", async () => {
    const up = await startFakeUpstream({
      status: 401,
      body: JSON.stringify({ error: { message: "invalid_api_key", type: "auth" } }),
    });
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ model: "x", stream: false }),
      });
      assert.equal(res.status, 401);
      assert.equal(res.json().error.message, "invalid_api_key");
      assert.equal(up.requests[0].headers.authorization, "Bearer test-secret-key");
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("maps upstream network failure to 502 fixed error", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({
      upstreamPort: up.port,
      fetchImpl: async () => {
        throw Object.assign(new Error("connect ECONNREFUSED"), { code: "ECONNREFUSED" });
      },
    });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(res.status, 502);
      assert.deepEqual(res.json(), { error: "upstream_error" });
      assert.equal(res.text.includes("ECONNREFUSED"), false);
      assert.equal(res.text.includes("test-secret-key"), false);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("streams SSE across odd chunk boundaries transparently", async () => {
    const full = "data: {\"delta\":\"hello-world\"}\n\ndata: [DONE]\n\n";
    const chunks = [];
    for (let i = 0; i < full.length; i += 3) {
      chunks.push(full.slice(i, i + 3));
    }
    const up = await startFakeUpstream({
      headers: { "content-type": "text/event-stream; charset=utf-8" },
      streamChunks: chunks,
    });
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      const res = await fetch(`${gw.base}/api/ai/v1/fake/responses`, {
        method: "POST",
        headers: {
          "content-type": "application/json; charset=utf-8",
          accept: "text/event-stream",
        },
        body: JSON.stringify({ model: "x", stream: true }),
      });
      assert.equal(res.status, 200);
      assert.match(res.headers.get("content-type") || "", /text\/event-stream/);
      assert.equal(res.headers.get("x-accel-buffering"), "no");
      assert.match(res.headers.get("cache-control") || "", /no-store/);
      const text = await res.text();
      assert.equal(text, full);
      assert.equal(up.requests[0].url, "/v1/responses");
      assert.equal(up.requests[0].headers.accept, "text/event-stream");
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("accepts application/json charset and anthropic messages path", async () => {
    const up = await startFakeUpstream({
      body: JSON.stringify({ type: "message", id: "msg_1" }),
    });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        ANTHROPIC_API_KEY: "ant-secret-value",
        AI_GATEWAY_UPSTREAMS_JSON: JSON.stringify({
          fake: {
            baseUrl: `http://127.0.0.1:${up.port}/v1`,
            auth: { type: "x-api-key", env: "ANTHROPIC_API_KEY" },
          },
        }),
      },
    });
    try {
      const res = await request(`${gw.base}/api/ai/v1/fake/messages`, {
        method: "POST",
        headers: {
          "content-type": "application/json; charset=utf-8",
          "anthropic-version": "2023-06-01",
          "anthropic-beta": "tools-2024-04-04",
          "x-api-key": "client-must-not-win",
          authorization: "Bearer client-must-not-win",
        },
        body: JSON.stringify({
          model: "claude",
          max_tokens: 16,
          messages: [{ role: "user", content: "hi" }],
          stream: true,
        }),
      });
      assert.equal(res.status, 200);
      assert.equal(res.json().id, "msg_1");
      const hit = up.requests[0];
      assert.equal(hit.url, "/v1/messages");
      assert.equal(hit.headers["x-api-key"], "ant-secret-value");
      assert.equal(hit.headers.authorization, undefined);
      assert.equal(hit.headers["anthropic-version"], "2023-06-01");
      assert.equal(hit.headers["anthropic-beta"], "tools-2024-04-04");
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("rejects open-proxy style paths and disallowed methods", async () => {
    const up = await startFakeUpstream();
    const gw = await startGateway({ upstreamPort: up.port });
    try {
      for (const path of [
        "/api/ai/v1/fake/chat/completions/../../models",
        "/api/ai/v1/fake//models",
        "/api/ai/v1/fake/http://evil.example/",
        "/api/ai/v1/fake/models/extra",
      ]) {
        const res = await request(`${gw.base}${path}`, {
          method: path.includes("chat") ? "POST" : "GET",
          headers: { "content-type": "application/json" },
          body: path.includes("chat") ? "{}" : undefined,
        });
        assert.equal(res.status, 404, path);
      }
      const put = await request(`${gw.base}/api/ai/v1/fake/models`, { method: "PUT" });
      assert.equal(put.status, 404);
      assert.equal(up.requests.length, 0);
    } finally {
      await gw.close();
      await up.close();
    }
  });

  it("gateway error responses and logs never leak API keys", async () => {
    const secret = "super-secret-key-do-not-leak";
    const logs = [];
    const originalError = console.error;
    const originalLog = console.log;
    console.error = (...args) => {
      logs.push(args.map(String).join(" "));
    };
    console.log = (...args) => {
      logs.push(args.map(String).join(" "));
    };

    const up = await startFakeUpstream({ hang: true });
    const gw = await startGateway({
      upstreamPort: up.port,
      env: {
        FAKE_API_KEY: secret,
        AI_GATEWAY_CONNECT_TIMEOUT_MS: "50",
        AI_GATEWAY_TOTAL_TIMEOUT_MS: "80",
      },
    });
    try {
      const timedOut = await request(`${gw.base}/api/ai/v1/fake/models`);
      assert.equal(timedOut.status, 504);
      assert.deepEqual(timedOut.json(), { error: "upstream_timeout" });
      assert.equal(timedOut.text.includes(secret), false);

      const unknown = await request(`${gw.base}/api/ai/v1/missing/models`);
      assert.equal(unknown.status, 404);
      assert.equal(unknown.text.includes(secret), false);

      const tooBig = await fetch(`${gw.base}/api/ai/v1/fake/chat/completions`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: Buffer.alloc(4 * 1024 * 1024 + 8, 0x7b),
      });
      assert.equal(tooBig.status, 413);
      const tooBigText = await tooBig.text();
      assert.equal(tooBigText.includes(secret), false);

      const joined = logs.join("\n");
      assert.equal(joined.includes(secret), false);
      assert.equal(joined.includes("Bearer "), false);
    } finally {
      console.error = originalError;
      console.log = originalLog;
      await gw.close();
      await up.close();
    }
  });
});
