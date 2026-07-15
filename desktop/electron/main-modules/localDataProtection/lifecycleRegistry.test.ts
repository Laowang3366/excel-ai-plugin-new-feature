import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import type { DataKeystoreCipher } from "./dataKeystore";
import {
  clearLocalDataProtectionRuntimeForTests,
  initializeLocalDataProtection,
  getLocalDataProtectionRuntime,
} from "./localDataProtectionService";
import { afterDataPathMigrated } from "./localDataLifecycle";
import { clearManagedDataRoot } from "./managedRootCleanup";
import { setPayloadProtection } from "./payloadProtection";

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
  clearLocalDataProtectionRuntimeForTests();
  setPayloadProtection(null);
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe("data path migration old-root registry", () => {
  it("clears old root managed data and settings on success", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "mig-old-ok-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const oldRoot = path.join(base, "old");
    const nextRoot = path.join(base, "next");
    fs.mkdirSync(path.join(oldRoot, "sessions"), { recursive: true });
    fs.mkdirSync(path.join(oldRoot, "settings"), { recursive: true });
    fs.writeFileSync(path.join(oldRoot, "sessions", "a.txt"), "old");
    fs.writeFileSync(path.join(oldRoot, "settings", "s.json"), "{}");
    fs.mkdirSync(nextRoot, { recursive: true });

    initializeLocalDataProtection({
      dataRoot: nextRoot,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    const result = await afterDataPathMigrated({
      previousDataPath: oldRoot,
      nextDataPath: nextRoot,
    });
    expect(result.oldRootCleared).toBe(true);
    expect(fs.existsSync(path.join(oldRoot, "sessions"))).toBe(false);
    expect(fs.existsSync(path.join(oldRoot, "settings"))).toBe(false);
    const erasable = getLocalDataProtectionRuntime()!.registry.listErasable();
    expect(erasable.some((e) => e.absolutePath === oldRoot && e.status === "pending_erase")).toBe(
      false,
    );
  });

  it("returns oldRootError and keeps pending_erase when cleanup fails", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "mig-old-fail-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const oldRoot = path.join(base, "old");
    const nextRoot = path.join(base, "next");
    fs.mkdirSync(oldRoot, { recursive: true });
    // Make sessions a file so eraseManagedUserData fails that category
    fs.writeFileSync(path.join(oldRoot, "sessions"), "not-dir");
    fs.mkdirSync(nextRoot, { recursive: true });

    initializeLocalDataProtection({
      dataRoot: nextRoot,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    // afterDataPathMigrated registers old_root then clears
    const result = await afterDataPathMigrated({
      previousDataPath: oldRoot,
      nextDataPath: nextRoot,
    });
    expect(result.oldRootCleared).toBe(false);
    expect(result.oldRootError).toBeTruthy();
    const pending = getLocalDataProtectionRuntime()!.registry.listErasable();
    expect(
      pending.some(
        (e) =>
          path.resolve(e.absolutePath) === path.resolve(oldRoot) && e.status === "pending_erase",
      ),
    ).toBe(true);
  });
});

describe("export registration on restore failure", () => {
  it("registers committed exportPath even when success is false", async () => {
    const base = fs.mkdtempSync(path.join(os.tmpdir(), "exp-reg-"));
    tempDirs.push(base);
    const userData = path.join(base, "userData");
    const dataRoot = path.join(base, "data");
    const exportTarget = path.join(base, "export-target");
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    fs.writeFileSync(path.join(dataRoot, "sessions", "s.txt"), "s");
    fs.mkdirSync(exportTarget, { recursive: true });

    initializeLocalDataProtection({
      dataRoot,
      userDataPath: userData,
      cipher: makeCipher(),
    });

    const { afterUserDataExported } = await import("./localDataLifecycle");
    // Simulate coordinator outcome: export committed, restore failed.
    const committedPath = path.join(exportTarget, "committed");
    fs.mkdirSync(committedPath, { recursive: true });
    fs.writeFileSync(path.join(committedPath, "privacy-export-manifest.json"), "{}");
    const result = {
      success: false as const,
      exportPath: committedPath,
      error: "数据已导出，但恢复本地运行时失败：restore failed",
    };
    if (result.exportPath) {
      afterUserDataExported(result.exportPath);
    }

    expect(result.success).toBe(false);
    expect(result.exportPath).toBeTruthy();
    expect(fs.existsSync(result.exportPath)).toBe(true);
    const registered = getLocalDataProtectionRuntime()!.registry.list();
    expect(
      registered.some(
        (e) =>
          path.resolve(e.absolutePath) === path.resolve(result.exportPath) &&
          e.category === "privacy_export",
      ),
    ).toBe(true);
  });
});
