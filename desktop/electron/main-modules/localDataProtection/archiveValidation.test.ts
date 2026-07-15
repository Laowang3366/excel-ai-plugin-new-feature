import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import * as zlib from "node:zlib";
import { promisify } from "node:util";
import { afterEach, describe, expect, it } from "vitest";
import { openOrCreateDataKeystore, type DataKeystoreCipher } from "./dataKeystore";
import { migrateLocalDataAtomically } from "./localDataMigrator";
import { createPayloadProtection, setPayloadProtection } from "./payloadProtection";

const gzipAsync = promisify(zlib.gzip);
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

describe("archive validation blocks incomplete encryption", () => {
  it("rejects migration when .jsonl.gz contains invalid ciphertext", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-arch-"));
    tempDirs.push(root);
    const dataRoot = path.join(root, "data");
    const userData = path.join(root, "userData");
    const sessions = path.join(dataRoot, "sessions");
    fs.mkdirSync(sessions, { recursive: true });
    await fs.promises.writeFile(
      path.join(sessions, "rollout-thread-z.jsonl.gz"),
      await gzipAsync(Buffer.from("ldp:v1:not-valid-base64!!!\n", "utf8")),
    );

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    const protection = createPayloadProtection(keystore);
    setPayloadProtection(protection);

    await expect(
      migrateLocalDataAtomically({
        dataRoot,
        protection,
        targetKeyId: 1,
        previousKeyId: 1,
        kind: "encrypt",
        userDataPath: userData,
      }),
    ).rejects.toThrow();
    // data root must remain (no successful swap)
    expect(fs.existsSync(dataRoot)).toBe(true);
  });

  it("accepts plaintext .jsonl.gz by sealing then validating decryptability", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-arch-ok-"));
    tempDirs.push(root);
    const dataRoot = path.join(root, "data");
    const userData = path.join(root, "userData");
    const sessions = path.join(dataRoot, "sessions");
    fs.mkdirSync(sessions, { recursive: true });
    const plainLine = `${JSON.stringify({
      timestamp: "t",
      item: { type: "user_message", content: "ARCH-OK" },
    })}\n`;
    await fs.promises.writeFile(
      path.join(sessions, "rollout-thread-ok.jsonl.gz"),
      await gzipAsync(Buffer.from(plainLine, "utf8")),
    );

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    const protection = createPayloadProtection(keystore);
    setPayloadProtection(protection);

    const handle = await migrateLocalDataAtomically({
      dataRoot,
      protection,
      targetKeyId: 1,
      previousKeyId: 1,
      kind: "encrypt",
      userDataPath: userData,
    });
    expect("finalize" in handle).toBe(true);
    if ("finalize" in handle) {
      const finalized = await handle.finalize();
      expect(finalized.cleared).toBe(true);
    }
  });
});
