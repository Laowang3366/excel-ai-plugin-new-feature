import { generateKeyPairSync, sign } from "node:crypto";

import { describe, expect, it } from "vitest";

import {
  canonicalJson,
  compareVersions,
  verifyRemoteUpdateManifest,
} from "./updateManifest";

function createSignedManifest() {
  const { publicKey, privateKey } = generateKeyPairSync("ed25519");
  const unsigned = {
    schemaVersion: 1 as const,
    channel: "stable" as const,
    version: "0.1.79",
    publishedAt: "2026-07-12T06:00:00.000Z",
    releaseNotes: ["新增应用内更新"],
    installer: {
      url: "https://plugin.shelelove.top/releases/windows/app.exe",
      sha256: "a".repeat(64),
      size: 1024,
    },
  };
  return {
    publicKey: publicKey.export({ type: "spki", format: "pem" }),
    manifest: {
      ...unsigned,
      signature: sign(null, Buffer.from(canonicalJson(unsigned)), privateKey).toString("base64"),
    },
  };
}

describe("verifyRemoteUpdateManifest", () => {
  it("accepts a correctly signed manifest", () => {
    const { manifest, publicKey } = createSignedManifest();
    expect(verifyRemoteUpdateManifest(manifest, publicKey).version).toBe("0.1.79");
  });

  it("rejects changes made after signing", () => {
    const { manifest, publicKey } = createSignedManifest();
    expect(() => verifyRemoteUpdateManifest({ ...manifest, version: "0.1.80" }, publicKey))
      .toThrow("更新清单签名无效");
  });
});

describe("compareVersions", () => {
  it("orders stable semantic versions", () => {
    expect(compareVersions("0.1.79", "0.1.78")).toBe(1);
    expect(compareVersions("0.1.78", "0.1.78")).toBe(0);
    expect(compareVersions("0.1.78-beta.1", "0.1.78")).toBe(-1);
  });
});
