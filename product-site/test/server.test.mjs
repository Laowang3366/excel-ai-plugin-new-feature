import assert from "node:assert/strict";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { hashPassword, verifyPassword } from "../src/auth.mjs";
import { buildServer } from "../src/server.mjs";

test("password hashes verify without storing plaintext", async () => {
  const encoded = await hashPassword("strong-test-password");
  assert.equal(await verifyPassword("strong-test-password", encoded), true);
  assert.equal(await verifyPassword("wrong-password", encoded), false);
});

test("download tracking and admin authentication work end to end", async (context) => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "wenge-site-"));
  const releasesDir = path.join(root, "releases");
  await fs.mkdir(releasesDir, { recursive: true });
  const artifact = "Wengge-AI-Assistant-Setup-0.1.79.exe";
  await fs.writeFile(path.join(releasesDir, artifact), "installer-bytes");
  await fs.writeFile(path.join(releasesDir, "release.json"), JSON.stringify({
    version: "0.1.79",
    publishedAt: "2026-07-12T06:00:00.000Z",
    releaseNotes: ["新增应用内更新"],
    installer: { fileName: artifact, size: 15, sha256: "a".repeat(64), downloadUrl: "/download/windows" },
  }));
  await fs.writeFile(path.join(releasesDir, "manifest.json"), JSON.stringify({ signed: true }));
  const password = "strong-test-password";
  const { app } = await buildServer({
    logger: false,
    dataDir: path.join(root, "data"),
    databasePath: path.join(root, "data", "analytics.sqlite"),
    releasesDir,
    adminPasswordHash: await hashPassword(password),
    cookieSecret: "test-cookie-secret-with-at-least-32-characters",
    analyticsSalt: "test-analytics-salt",
    publicDir: path.resolve(import.meta.dirname, "../public"),
  });
  context.after(async () => {
    await app.close();
    await fs.rm(root, { recursive: true, force: true });
  });

  assert.equal((await app.inject({ method: "GET", url: "/healthz" })).statusCode, 200);
  assert.equal((await app.inject({ method: "GET", url: "/api/admin/stats" })).statusCode, 401);

  const download = await app.inject({
    method: "GET",
    url: "/download/windows",
    headers: { "user-agent": "test-client", referer: "https://example.test" },
  });
  assert.equal(download.statusCode, 200);
  assert.equal(download.body, "installer-bytes");

  const login = await app.inject({
    method: "POST",
    url: "/api/admin/login",
    payload: { password },
  });
  assert.equal(login.statusCode, 200);
  const cookie = login.cookies.find((item) => item.name === "wenge_admin");
  assert.ok(cookie);

  const stats = await app.inject({
    method: "GET",
    url: "/api/admin/stats?days=30",
    cookies: { wenge_admin: cookie.value },
  });
  assert.equal(stats.statusCode, 200);
  assert.equal(stats.json().summary.total, 1);
  assert.equal(stats.json().recent[0].userAgent, "test-client");
});
