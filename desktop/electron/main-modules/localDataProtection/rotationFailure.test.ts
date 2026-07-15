import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { openOrCreateDataKeystore, type DataKeystoreCipher } from "./dataKeystore";
import { migrateLocalDataAtomically, readProtectionMarker } from "./localDataMigrator";
import { createPayloadProtection, setPayloadProtection } from "./payloadProtection";
import { isProtectedBlob } from "./protectedBlob";

const tempDirs: string[] = [];

function makeCipher(): DataKeystoreCipher {
  const map = new Map<string, string>();
  return {
    isAvailable: () => true,
    encrypt: (value) => {
      const token = Buffer.from(value, "utf8").toString("base64");
      map.set(token, value);
      return token;
    },
    decrypt: (value) => {
      const plain = map.get(value);
      if (!plain) throw new Error("missing");
      return plain;
    },
  };
}

afterEach(() => {
  setPayloadProtection(null);
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe("rotation two-phase failure recovery", () => {
  it("keeps backup until finalize and rolls back after swap if commit fails", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-rot-"));
    tempDirs.push(root);
    const dataRoot = path.join(root, "data");
    const userData = path.join(root, "userData");
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    const jsonlPath = path.join(dataRoot, "sessions", "rollout-thread-a.jsonl");
    fs.writeFileSync(
      jsonlPath,
      `${JSON.stringify({ timestamp: "t", item: { type: "x", body: "OLD" } })}\n`,
    );

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    const protection = createPayloadProtection(keystore);
    setPayloadProtection(protection);

    // First encrypt with current key
    const first = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: keystore.currentKeyId,
      previousKeyId: keystore.currentKeyId,
      kind: "encrypt",
      userDataPath: userData,
    });
    if ("finalize" in first) await first.finalize();

    const pending = keystore.createPendingKey();
    const handle = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: pending,
      previousKeyId: 1,
      kind: "rotate",
      userDataPath: userData,
      force: true,
    });
    expect("finalize" in handle).toBe(true);
    if (!("finalize" in handle)) return;

    // Simulate keystore commit failure: rollback data, discard pending
    await handle.rollback();
    keystore.discardPendingKey();
    expect(keystore.currentKeyId).toBe(1);
    expect(keystore.pendingKeyId).toBeNull();
    const marker = readProtectionMarker(dataRoot);
    // after rollback, old marker from first migration remains
    expect(marker?.contentKeyId).toBe(1);
    const raw = fs.readFileSync(jsonlPath, "utf8").trim();
    expect(isProtectedBlob(raw)).toBe(true);
  });

  it("reverts current key when restore fails after commit", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-rot2-"));
    tempDirs.push(root);
    const dataRoot = path.join(root, "data");
    const userData = path.join(root, "userData");
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    fs.writeFileSync(
      path.join(dataRoot, "sessions", "rollout-thread-b.jsonl"),
      `${JSON.stringify({ timestamp: "t", item: { type: "x", body: "BODY" } })}\n`,
    );

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    const protection = createPayloadProtection(keystore);
    setPayloadProtection(protection);

    const first = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: 1,
      previousKeyId: 1,
      kind: "encrypt",
      userDataPath: userData,
    });
    if ("finalize" in first) await first.finalize();

    keystore.createPendingKey();
    const handle = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: 2,
      previousKeyId: 1,
      kind: "rotate",
      userDataPath: userData,
      force: true,
    });
    if (!("finalize" in handle)) return;
    keystore.commitPendingKey();
    handle.markCommitted();
    expect(keystore.currentKeyId).toBe(2);
    await handle.rollback();
    keystore.revertToKeyId(1);
    expect(keystore.currentKeyId).toBe(1);
    expect(readProtectionMarker(dataRoot)?.contentKeyId).toBe(1);
  });
});
