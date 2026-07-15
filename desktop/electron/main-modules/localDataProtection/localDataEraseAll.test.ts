import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import type { DataKeystoreCipher } from "./dataKeystore";
import {
  clearLocalDataProtectionRuntimeForTests,
  initializeLocalDataProtection,
  registerManagedDataRoot,
  getLocalDataProtectionRuntime,
} from "./localDataProtectionService";
import { eraseAllManagedReplicasAndKeys } from "./localDataEraseAll";
import { getPayloadProtection, setPayloadProtection } from "./payloadProtection";
import { USER_DATA_ERASE_CONFIRMATION } from "../userDataErase";

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

function makeDataRoot(base: string, name: string): string {
  const root = path.join(base, name);
  fs.mkdirSync(path.join(root, "sessions"), { recursive: true });
  fs.mkdirSync(path.join(root, "settings"), { recursive: true });
  fs.writeFileSync(path.join(root, "sessions", "note.txt"), "session-data");
  fs.writeFileSync(path.join(root, "settings", "excel-ai-settings.json"), "{}");
  return root;
}

afterEach(() => {
  clearLocalDataProtectionRuntimeForTests();
  setPayloadProtection(null);
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe("eraseAllManagedReplicasAndKeys", () => {
  it("rejects parent/child registered paths before any delete or key destroy", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "erase-nest-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const parent = makeDataRoot(base, "parent");
    const child = path.join(parent, "child-export");
    fs.mkdirSync(child, { recursive: true });
    fs.writeFileSync(path.join(child, "x.txt"), "x");

    const runtime = initializeLocalDataProtection({
      dataRoot: parent,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    registerManagedDataRoot(child, "privacy_export");
    const installBefore = runtime.keystore.installId;
    const keyIdsBefore = runtime.keystore.listKeyIds();

    const result = await eraseAllManagedReplicasAndKeys({ userDataPath: userData });
    expect(result.success).toBe(false);
    expect(result.keyMaterialDestroyed).toBe(false);
    expect(result.errors.some((e) => e.includes("根包含关系"))).toBe(true);
    expect(fs.existsSync(path.join(parent, "sessions", "note.txt"))).toBe(true);
    expect(fs.existsSync(path.join(child, "x.txt"))).toBe(true);
    expect(runtime.keystore.installId).toBe(installBefore);
    expect(runtime.keystore.listKeyIds()).toEqual(keyIdsBefore);
    expect(result.proofSummary?.destroyedKeyCount).toBe(0);
    expect(result.proofSummary?.keyMaterialDestroyed).toBe(false);
  });

  it("fails before any delete when replacement pending key cannot be created", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "erase-repkey-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const active = makeDataRoot(base, "active");
    const cipher = makeCipher();

    const runtime = initializeLocalDataProtection({
      dataRoot: active,
      userDataPath: userData,
      cipher,
    });
    const keysBefore = runtime.keystore.listKeyIds();
    const currentBefore = runtime.keystore.currentKeyId;
    const protection = getPayloadProtection();
    expect(protection).toBeTruthy();
    const sealed = protection!.protect("canary-plain", "aad:test");

    cipher.isAvailable = () => false;

    const result = await eraseAllManagedReplicasAndKeys({ userDataPath: userData });
    expect(result.success).toBe(false);
    expect(result.keyMaterialDestroyed).toBe(false);
    expect(result.errors.some((e) => e.includes("replacement_key"))).toBe(true);
    expect(fs.existsSync(path.join(active, "sessions", "note.txt"))).toBe(true);
    expect(runtime.keystore.listKeyIds()).toEqual(keysBefore);
    expect(runtime.keystore.currentKeyId).toBe(currentBefore);

    cipher.isAvailable = () => true;
    const still = getPayloadProtection();
    expect(still).toBeTruthy();
    expect(still!.unprotect(sealed, "aad:test")).toBe("canary-plain");
    expect(still!.protect("post-fail", "aad:test").length).toBeGreaterThan(0);
  });

  it("removes entire privacy_export root and writes accurate proof json", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "erase-ok-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const active = makeDataRoot(base, "active");
    const exportDir = path.join(base, "export-copy");
    fs.mkdirSync(path.join(exportDir, "sessions"), { recursive: true });
    fs.mkdirSync(path.join(exportDir, "settings"), { recursive: true });
    fs.writeFileSync(path.join(exportDir, "privacy-export-manifest.json"), "{}");
    fs.writeFileSync(path.join(exportDir, "settings", "privacy-export-settings.json"), "{}");
    fs.writeFileSync(path.join(exportDir, "sessions", "e.txt"), "export-data");

    const runtime = initializeLocalDataProtection({
      dataRoot: active,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    registerManagedDataRoot(exportDir, "privacy_export");
    const installBefore = runtime.keystore.installId;
    // Simulate interrupted prior erase that left a pending replacement key.
    const leftoverPending = runtime.keystore.createPendingKey();
    const priorKeys = runtime.keystore.listKeyIds().filter((id) => id !== leftoverPending);

    const result = await eraseAllManagedReplicasAndKeys({ userDataPath: userData });
    expect(result.success).toBe(true);
    expect(result.keyMaterialDestroyed).toBe(true);
    expect(fs.existsSync(path.join(active, "sessions"))).toBe(false);
    expect(fs.existsSync(exportDir)).toBe(false);
    expect(runtime.keystore.installId).toBe(installBefore);
    // Reused leftover pending as sole replacement — prior keys purged
    expect(runtime.keystore.listKeyIds()).toEqual([leftoverPending]);
    expect(runtime.keystore.currentKeyId).toBe(leftoverPending);
    expect(priorKeys.every((id) => !runtime.keystore.listKeyIds().includes(id))).toBe(true);
    expect(getPayloadProtection()?.currentKeyId()).toBe(leftoverPending);

    const proofsDir = path.join(userData, "erase-proofs");
    const proofFiles = fs.readdirSync(proofsDir).filter((n) => n.endsWith(".json"));
    expect(proofFiles.length).toBeGreaterThan(0);
    const proofBody = fs.readFileSync(path.join(proofsDir, proofFiles[0]!), "utf8");
    expect(proofBody).not.toContain(active);
    expect(proofBody).not.toContain("export-data");
    expect(proofBody).toContain("keyMaterialDestroyed");
    expect(proofBody).toContain('"keyMaterialDestroyed": true');
    expect(proofBody).toContain(result.proofSummary!.installIdDigest);
    const parsed = JSON.parse(proofBody) as {
      keyDestruction: { keyMaterialDestroyed: boolean; destroyedKeyIds: number[] };
      replicas: Array<{ pathDigest: string; category: string; status: string }>;
    };
    expect(parsed.keyDestruction.keyMaterialDestroyed).toBe(true);
    expect(parsed.keyDestruction.destroyedKeyIds).toEqual(priorKeys);
    expect(parsed.keyDestruction.destroyedKeyIds).not.toContain(leftoverPending);
    expect(parsed.replicas.every((r) => r.status === "erased" || r.status === "failed")).toBe(true);
  });

  it("keeps pending_erase and records key destroy without inventing keystoreRemoved", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "erase-partial-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const active = makeDataRoot(base, "active");
    fs.rmSync(path.join(active, "sessions"), { recursive: true, force: true });
    fs.writeFileSync(path.join(active, "sessions"), "not-a-dir");
    const exportDir = path.join(base, "export-ok");
    fs.mkdirSync(path.join(exportDir, "sessions"), { recursive: true });
    fs.writeFileSync(path.join(exportDir, "sessions", "e.txt"), "e");

    const runtime = initializeLocalDataProtection({
      dataRoot: active,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    registerManagedDataRoot(exportDir, "privacy_export");
    const installBefore = runtime.keystore.installId;
    const priorKeys = runtime.keystore.listKeyIds();

    const result = await eraseAllManagedReplicasAndKeys({ userDataPath: userData });
    expect(result.success).toBe(false);
    expect(runtime.keystore.installId).toBe(installBefore);
    const pending = runtime.registry.listErasable();
    expect(pending.some((e) => e.status === "pending_erase")).toBe(true);
    expect(result.keyMaterialDestroyed).toBe(true);
    expect(runtime.keystore.listKeyIds()).toHaveLength(1);
    expect(priorKeys.every((id) => !runtime.keystore.listKeyIds().includes(id))).toBe(true);
    const proofBody = fs.readFileSync(result.proofPath!, "utf8");
    expect(proofBody).toContain('"keyMaterialDestroyed": true');
    expect(proofBody).not.toContain("keystoreRemoved");
  });
});

describe("runUserDataErase quiesced erase path", () => {
  it("does not swap keys when containment precheck fails", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "erase-hook-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const parent = makeDataRoot(base, "parent");
    const child = path.join(parent, "child");
    fs.mkdirSync(child, { recursive: true });

    const runtime = initializeLocalDataProtection({
      dataRoot: parent,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    registerManagedDataRoot(child, "privacy_export");
    const keysBefore = runtime.keystore.listKeyIds();
    let busyReleased = false;

    const { runEraseUserData } = await import("../settingsUserDataActions");
    const result = await runEraseUserData(USER_DATA_ERASE_CONFIRMATION, {
      isBusy: () => false,
      setBusy: (busy) => {
        if (!busy) busyReleased = true;
      },
      getAgents: () => [],
      getSessionStore: () =>
        ({
          suspendWrites: () => undefined,
          resumeWrites: () => undefined,
          flushRolloutWrites: async () => undefined,
        }) as never,
      closeStateRuntime: async () => undefined,
      resetKnowledge: () => undefined,
      clearSettings: () => undefined,
      resetSessionStore: async () => undefined,
      getActiveAIConfig: () =>
        ({
          provider: "openai",
          model: "m",
          apiKey: "",
          contextWindowSize: 1000,
        }) as never,
      getRuntimeSettingValue: () => undefined,
      reloadKnowledge: async () => ({ store: {} }),
      getSanitizedSettings: () => ({}),
    });

    expect(result.success).toBe(false);
    expect(busyReleased).toBe(true);
    expect(runtime.keystore.listKeyIds()).toEqual(keysBefore);
    expect(fs.existsSync(path.join(parent, "sessions", "note.txt"))).toBe(true);
  });
});
