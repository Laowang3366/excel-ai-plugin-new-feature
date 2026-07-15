import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import {
  openOrCreateDataKeystore,
  type DataKeystoreCipher,
} from "../../main-modules/localDataProtection/dataKeystore";
import {
  createPayloadProtection,
  setPayloadProtection,
} from "../../main-modules/localDataProtection/payloadProtection";
import {
  createRecordId,
  jsonlLineAad,
  sealUtf8,
} from "../../main-modules/localDataProtection/protectedBlob";
import { SessionStore } from "./sessionStore";

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

describe("SessionStore protected JSONL integrity", () => {
  it("rejects tampered ciphertext on load instead of silently skipping", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "session-cipher-"));
    tempDirs.push(root);
    const userData = path.join(root, "userData");
    const dataRoot = path.join(root, "data");
    const sessionsRoot = path.join(root, "sessions");
    fs.mkdirSync(userData, { recursive: true });
    fs.mkdirSync(dataRoot, { recursive: true });
    fs.mkdirSync(sessionsRoot, { recursive: true });

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    setPayloadProtection(createPayloadProtection(keystore));

    const store = new SessionStore(sessionsRoot);
    const thread = await store.createThread("provider", "model");
    await store.flushRolloutWrites();

    const files = fs.readdirSync(sessionsRoot, { recursive: true }) as string[];
    const jsonlRel = files.find((name) => String(name).endsWith(".jsonl"));
    expect(jsonlRel).toBeTruthy();
    const jsonlPath = path.join(sessionsRoot, String(jsonlRel));
    const original = fs.readFileSync(jsonlPath, "utf8").trim();
    // Tamper envelope payload while keeping prefix
    const tampered = `${original.slice(0, -8)}AAAAAAAA`;
    fs.writeFileSync(jsonlPath, `${tampered}\n`, "utf8");

    await expect(store.loadThread(thread.metadata.threadId)).rejects.toThrow();
  });

  it("still skips corrupt plaintext lines", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "session-plain-"));
    tempDirs.push(root);
    const sessionsRoot = path.join(root, "sessions");
    fs.mkdirSync(sessionsRoot, { recursive: true });
    const filePath = path.join(sessionsRoot, "rollout-thread-x.jsonl");
    fs.writeFileSync(
      filePath,
      [
        JSON.stringify({
          timestamp: "t",
          item: {
            type: "session_meta",
            meta: {
              id: "thread-x",
              timestamp: new Date().toISOString(),
              modelProvider: "p",
            },
          },
        }),
        "{not-json",
      ].join("\n"),
      "utf8",
    );
    const store = new SessionStore(sessionsRoot);
    const loaded = await store.loadThreadByPath(filePath);
    expect(loaded?.metadata.threadId).toBe("thread-x");
  });
});
