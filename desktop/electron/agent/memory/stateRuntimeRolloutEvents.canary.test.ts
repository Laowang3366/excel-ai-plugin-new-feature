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
import { StateRuntimeStore } from "./stateRuntimeStore";

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

describe("rollout ciphertext-first insert", () => {
  it("never leaves canary plaintext in logs.db or WAL after append", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "rollout-canary-"));
    tempDirs.push(root);
    const userData = path.join(root, "userData");
    const dataRoot = path.join(root, "data");
    const runtimeRoot = path.join(root, "state-runtime");
    fs.mkdirSync(userData, { recursive: true });
    fs.mkdirSync(dataRoot, { recursive: true });

    const { keystore } = openOrCreateDataKeystore({
      userDataPath: userData,
      dataRoot,
      cipher: makeCipher(),
    });
    setPayloadProtection(createPayloadProtection(keystore));

    const canary = "CANARY-ROLLOUT-PLAINTEXT-NEVER-ON-DISK";
    const store = new StateRuntimeStore(runtimeRoot);
    await store.init();
    await store.appendRolloutItems("thread-canary", [
      {
        type: "turn_item",
        turnId: "turn-1",
        item: {
          type: "user_message",
          id: "msg-1",
          content: canary,
          timestamp: 1,
        },
      },
    ]);
    const dbPath = store.getDatabasePaths().logs;
    await store.close();

    const rawDb = fs.readFileSync(dbPath);
    expect(rawDb.includes(Buffer.from(canary, "utf8"))).toBe(false);
    const walPath = `${dbPath}-wal`;
    if (fs.existsSync(walPath)) {
      const rawWal = fs.readFileSync(walPath);
      expect(rawWal.includes(Buffer.from(canary, "utf8"))).toBe(false);
    }

    // Readable after decrypt
    const reopened = new StateRuntimeStore(runtimeRoot);
    await reopened.init();
    const events = await reopened.listRolloutEvents("thread-canary");
    expect(JSON.stringify(events)).toContain(canary);
    await reopened.close();
  });
});
