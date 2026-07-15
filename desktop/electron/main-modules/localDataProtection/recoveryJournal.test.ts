import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import type { DataKeystoreCipher } from "./dataKeystore";
import { setPayloadProtection } from "./payloadProtection";
import { fixedTransactionPaths, writeMigrationJournal } from "./migrationJournal";
import { readProtectionMarker } from "./localDataMigrator";
import {
  clearLocalDataProtectionRuntimeForTests,
  initializeLocalDataProtection,
  recoverInterruptedMigration,
} from "./localDataProtectionService";

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

function setupRoots() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "ldp-rec-"));
  tempDirs.push(root);
  const dataRoot = path.join(root, "data");
  const userData = path.join(root, "userData");
  fs.mkdirSync(dataRoot, { recursive: true });
  fs.mkdirSync(userData, { recursive: true });
  return { dataRoot, userData };
}

describe("recoverInterruptedMigration journal guards", () => {
  it("restores backup on swapping rename-gap (data missing, backup+stage present)", () => {
    const { dataRoot, userData } = setupRoots();
    const cipher = makeCipher();
    initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher });

    const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
    fs.mkdirSync(backupRoot, { recursive: true });
    fs.writeFileSync(path.join(backupRoot, "keep.txt"), "old-root");
    fs.mkdirSync(stageRoot, { recursive: true });
    fs.writeFileSync(
      path.join(userData, "local-data-migration-journal.json"),
      JSON.stringify({
        formatVersion: 1,
        phase: "swapping",
        kind: "encrypt",
        dataRoot,
        stageRoot,
        backupRoot,
        targetKeyId: 1,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      }),
    );
    fs.rmSync(dataRoot, { recursive: true, force: true });

    recoverInterruptedMigration({ dataRoot, userDataPath: userData });
    expect(fs.existsSync(path.join(dataRoot, "keep.txt"))).toBe(true);
    expect(fs.existsSync(backupRoot)).toBe(false);
    expect(fs.existsSync(stageRoot)).toBe(false);
  });

  it("treats swapping with target live+backup as swapped and commits pending", () => {
    const { dataRoot, userData } = setupRoots();
    const cipher = makeCipher();
    const runtime = initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher });

    // Create pending key as if mid-rotation (must use runtime keystore)
    const pending = runtime.keystore.createPendingKey();
    const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    fs.writeFileSync(
      path.join(dataRoot, "sessions", ".local-data-protection.json"),
      JSON.stringify({
        formatVersion: 1,
        contentKeyId: pending,
        migratedAt: new Date().toISOString(),
      }),
    );
    fs.mkdirSync(backupRoot, { recursive: true });
    fs.writeFileSync(path.join(backupRoot, "old.txt"), "backup");
    // stage already consumed by second rename
    writeMigrationJournal(
      {
        formatVersion: 1,
        phase: "swapping",
        kind: "rotate",
        dataRoot,
        stageRoot,
        backupRoot,
        targetKeyId: pending,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      },
      userData,
    );

    recoverInterruptedMigration({ dataRoot, userDataPath: userData });
    expect(runtime.keystore.currentKeyId).toBe(pending);
    expect(runtime.keystore.pendingKeyId).toBeNull();
    expect(readProtectionMarker(dataRoot)?.contentKeyId).toBe(pending);
    // backup cleaned or registered; prefer cleaned
    expect(fs.existsSync(backupRoot)).toBe(false);
  });

  it("purges retired keys when target already current after commit-before-crash", () => {
    const { dataRoot, userData } = setupRoots();
    const cipher = makeCipher();
    const runtime = initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher });
    const pending = runtime.keystore.createPendingKey();
    // Simulate: commitPendingKey succeeded, purgeRetiredKeys never ran.
    runtime.keystore.commitPendingKey();
    expect(runtime.keystore.currentKeyId).toBe(pending);
    expect(runtime.keystore.listKeyIds()).toContain(1);
    expect(runtime.keystore.listKeyIds()).toContain(pending);

    const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
    fs.mkdirSync(path.join(dataRoot, "sessions"), { recursive: true });
    fs.writeFileSync(
      path.join(dataRoot, "sessions", ".local-data-protection.json"),
      JSON.stringify({
        formatVersion: 1,
        contentKeyId: pending,
        migratedAt: new Date().toISOString(),
      }),
    );
    fs.mkdirSync(backupRoot, { recursive: true });
    fs.writeFileSync(path.join(backupRoot, "old.txt"), "backup");
    writeMigrationJournal(
      {
        formatVersion: 1,
        phase: "swapped",
        kind: "rotate",
        dataRoot,
        stageRoot,
        backupRoot,
        targetKeyId: pending,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      },
      userData,
    );

    recoverInterruptedMigration({ dataRoot, userDataPath: userData });
    expect(runtime.keystore.currentKeyId).toBe(pending);
    expect(runtime.keystore.listKeyIds()).toEqual([pending]);
    expect(fs.existsSync(backupRoot)).toBe(false);
    expect(fs.existsSync(path.join(userData, "local-data-migration-journal.json"))).toBe(false);
    expect(readProtectionMarker(dataRoot)?.contentKeyId).toBe(pending);
  });

  function writeEmptyKeystore(userData: string, installId: string): string {
    const keystorePath = path.join(userData, "local-data-keystore.json");
    fs.writeFileSync(
      keystorePath,
      `${JSON.stringify(
        {
          formatVersion: 1,
          installId,
          currentKeyId: 0,
          pendingKeyId: null,
          keys: [],
        },
        null,
        2,
      )}\n`,
      "utf8",
    );
    return keystorePath;
  }

  function writeMarker(root: string, keyId = 1): void {
    fs.mkdirSync(path.join(root, "sessions"), { recursive: true });
    fs.writeFileSync(
      path.join(root, "sessions", ".local-data-protection.json"),
      JSON.stringify({
        formatVersion: 1,
        contentKeyId: keyId,
        migratedAt: new Date().toISOString(),
      }),
    );
  }

  it("E clean first install creates keystore; clean legacy empty reseeds", () => {
    const { dataRoot, userData } = setupRoots();
    const keystorePath = path.join(userData, "local-data-keystore.json");
    expect(fs.existsSync(keystorePath)).toBe(false);
    const created = initializeLocalDataProtection({
      dataRoot,
      userDataPath: userData,
      cipher: makeCipher(),
    });
    expect(fs.existsSync(keystorePath)).toBe(true);
    expect(created.keystore.listKeyIds()).toEqual([1]);
    clearLocalDataProtectionRuntimeForTests();
    setPayloadProtection(null);

    const { dataRoot: data2, userData: user2 } = setupRoots();
    writeEmptyKeystore(user2, "clean-empty");
    const reseeds = initializeLocalDataProtection({
      dataRoot: data2,
      userDataPath: user2,
      cipher: makeCipher(),
    });
    expect(reseeds.keystore.installId).toBe("clean-empty");
    expect(reseeds.keystore.listKeyIds()).toEqual([1]);
  });

  it("A missing keystore + active marker fails and does not create keystore file", () => {
    const { dataRoot, userData } = setupRoots();
    const keystorePath = path.join(userData, "local-data-keystore.json");
    writeMarker(dataRoot);
    expect(() =>
      initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher: makeCipher() }),
    ).toThrow(/empty_keystore_with_protected_data/);
    expect(fs.existsSync(keystorePath)).toBe(false);
  });

  it("B missing keystore + dataRoot missing + fixed backup/stage markers fail without creating keystore", () => {
    const { dataRoot, userData } = setupRoots();
    const keystorePath = path.join(userData, "local-data-keystore.json");
    const { stageRoot, backupRoot } = fixedTransactionPaths(dataRoot);
    writeMarker(backupRoot);
    writeMarker(stageRoot, 2);
    fs.rmSync(dataRoot, { recursive: true, force: true });
    expect(() =>
      initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher: makeCipher() }),
    ).toThrow(/empty_keystore_with_protected_data/);
    expect(fs.existsSync(keystorePath)).toBe(false);
  });

  it("C empty keystore + untrusted journal fails and leaves keystore bytes unchanged", () => {
    const { dataRoot, userData } = setupRoots();
    const keystorePath = writeEmptyKeystore(userData, "blocked-untrusted");
    const before = fs.readFileSync(keystorePath, "utf8");
    writeMigrationJournal(
      {
        formatVersion: 1,
        phase: "swapping",
        kind: "encrypt",
        dataRoot: path.join(userData, "other-data"),
        stageRoot: path.join(userData, "other-stage"),
        backupRoot: path.join(userData, "evil-backup"),
        targetKeyId: 1,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      },
      userData,
    );
    expect(() =>
      initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher: makeCipher() }),
    ).toThrow(/empty_keystore_with_protected_data/);
    expect(fs.readFileSync(keystorePath, "utf8")).toBe(before);
  });

  it("D missing keystore + untrusted journal fails without creating keystore", () => {
    const { dataRoot, userData } = setupRoots();
    const keystorePath = path.join(userData, "local-data-keystore.json");
    writeMigrationJournal(
      {
        formatVersion: 1,
        phase: "swapping",
        kind: "encrypt",
        dataRoot: path.join(userData, "other-data"),
        stageRoot: path.join(userData, "other-stage"),
        backupRoot: path.join(userData, "evil-backup"),
        targetKeyId: 1,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      },
      userData,
    );
    expect(() =>
      initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher: makeCipher() }),
    ).toThrow(/empty_keystore_with_protected_data/);
    expect(fs.existsSync(keystorePath)).toBe(false);
  });

  it("rejects journal path tampering on recovery without following evil paths", () => {
    const { dataRoot, userData } = setupRoots();
    const cipher = makeCipher();
    initializeLocalDataProtection({ dataRoot, userDataPath: userData, cipher });
    const evil = path.join(userData, "evil-backup");
    fs.mkdirSync(evil, { recursive: true });
    fs.writeFileSync(path.join(evil, "x.txt"), "x");
    writeMigrationJournal(
      {
        formatVersion: 1,
        phase: "swapping",
        kind: "encrypt",
        dataRoot: path.join(userData, "other-data"),
        stageRoot: path.join(userData, "other-stage"),
        backupRoot: evil,
        targetKeyId: 1,
        previousKeyId: 1,
        updatedAt: new Date().toISOString(),
      },
      userData,
    );

    expect(() => recoverInterruptedMigration({ dataRoot, userDataPath: userData })).toThrow(
      /migration_journal_/,
    );
    expect(fs.existsSync(path.join(evil, "x.txt"))).toBe(true);
  });
});
