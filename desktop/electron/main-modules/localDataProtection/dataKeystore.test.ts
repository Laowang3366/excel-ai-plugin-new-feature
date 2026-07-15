import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { openOrCreateDataKeystore, type DataKeystoreCipher } from "./dataKeystore";

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
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe("DataKeystore pending rotation", () => {
  it("creates pending key without switching current until commit", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-keystore-"));
    tempDirs.push(dir);
    const dataRoot = path.join(dir, "data");
    fs.mkdirSync(dataRoot, { recursive: true });
    const { keystore } = openOrCreateDataKeystore({
      userDataPath: dir,
      dataRoot,
      cipher: makeCipher(),
    });
    expect(keystore.currentKeyId).toBe(1);
    const pending = keystore.createPendingKey();
    expect(pending).toBe(2);
    expect(keystore.currentKeyId).toBe(1);
    expect(keystore.pendingKeyId).toBe(2);
    keystore.getKey(2);
    const committed = keystore.commitPendingKey();
    expect(committed.nextKeyId).toBe(2);
    expect(keystore.currentKeyId).toBe(2);
    expect(keystore.listKeyIds()).toEqual([1, 2]);
    keystore.purgeRetiredKeys();
    expect(keystore.listKeyIds()).toEqual([2]);
  });

  it("discards pending key on failure path", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-keystore-"));
    tempDirs.push(dir);
    const dataRoot = path.join(dir, "data");
    fs.mkdirSync(dataRoot, { recursive: true });
    const { keystore } = openOrCreateDataKeystore({
      userDataPath: dir,
      dataRoot,
      cipher: makeCipher(),
    });
    keystore.createPendingKey();
    keystore.discardPendingKey();
    expect(keystore.pendingKeyId).toBeNull();
    expect(keystore.currentKeyId).toBe(1);
    expect(keystore.listKeyIds()).toEqual([1]);
  });

  it("atomically commits pending and purges all prior keys in one write", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-keystore-"));
    tempDirs.push(dir);
    const dataRoot = path.join(dir, "data");
    fs.mkdirSync(dataRoot, { recursive: true });
    const { keystore } = openOrCreateDataKeystore({
      userDataPath: dir,
      dataRoot,
      cipher: makeCipher(),
    });
    const prior = keystore.listKeyIds();
    const pending = keystore.createPendingKey();
    const swapped = keystore.commitPendingKeyAndPurgePriorKeys();
    expect(swapped.replacementKeyId).toBe(pending);
    expect(swapped.destroyedKeyIds).toEqual(prior);
    expect(keystore.currentKeyId).toBe(pending);
    expect(keystore.pendingKeyId).toBeNull();
    expect(keystore.listKeyIds()).toEqual([pending]);
    const onDisk = JSON.parse(
      fs.readFileSync(path.join(dir, "local-data-keystore.json"), "utf8"),
    ) as { currentKeyId: number; pendingKeyId: number | null; keys: Array<{ keyId: number }> };
    expect(onDisk.currentKeyId).toBe(pending);
    expect(onDisk.pendingKeyId).toBeNull();
    expect(onDisk.keys.map((k) => k.keyId)).toEqual([pending]);
  });

  it("reseeds empty keystore only when dataRoot has no protected artifacts", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-keystore-"));
    tempDirs.push(dir);
    const dataRoot = path.join(dir, "data");
    fs.mkdirSync(dataRoot, { recursive: true });
    const keystorePath = path.join(dir, "local-data-keystore.json");
    const emptyBody = `${JSON.stringify(
      {
        formatVersion: 1,
        installId: "install-empty-ok",
        currentKeyId: 0,
        pendingKeyId: null,
        keys: [],
      },
      null,
      2,
    )}\n`;
    fs.writeFileSync(keystorePath, emptyBody, "utf8");

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: dir,
      dataRoot,
      cipher: makeCipher(),
    });
    expect(keystore.currentKeyId).toBe(1);
    expect(keystore.listKeyIds()).toEqual([1]);
    expect(keystore.installId).toBe("install-empty-ok");
  });

  it("refuses to reseed empty keystore when marker present and leaves file unchanged", () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-keystore-"));
    tempDirs.push(dir);
    const dataRoot = path.join(dir, "data");
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    fs.writeFileSync(
      path.join(dataRoot, "sessions", ".local-data-protection.json"),
      JSON.stringify({ formatVersion: 1, contentKeyId: 1, migratedAt: new Date().toISOString() }),
    );
    const keystorePath = path.join(dir, "local-data-keystore.json");
    const emptyBody = `${JSON.stringify(
      {
        formatVersion: 1,
        installId: "install-empty-blocked",
        currentKeyId: 0,
        pendingKeyId: null,
        keys: [],
      },
      null,
      2,
    )}\n`;
    fs.writeFileSync(keystorePath, emptyBody, "utf8");
    const before = fs.readFileSync(keystorePath, "utf8");

    expect(() =>
      openOrCreateDataKeystore({
        userDataPath: dir,
        dataRoot,
        cipher: makeCipher(),
      }),
    ).toThrow(/empty_keystore_with_protected_data/);
    expect(fs.readFileSync(keystorePath, "utf8")).toBe(before);
  });
});
