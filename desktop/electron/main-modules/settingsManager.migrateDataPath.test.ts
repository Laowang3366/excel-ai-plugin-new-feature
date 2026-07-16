import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

const testState = vi.hoisted(() => {
  const stores = new Map<string, Record<string, unknown>>();
  return {
    stores,
    activeDataPath: "",
    reloadKnowledge: vi.fn(),
    resetKnowledge: vi.fn(),
    logMigrateFailure: vi.fn(),
  };
});

vi.mock("electron", () => ({
  safeStorage: {
    isEncryptionAvailable: () => false,
    encryptString: (value: string) => Buffer.from(value, "utf8"),
    decryptString: (value: Buffer) => value.toString("utf8"),
  },
  app: {
    isPackaged: false,
    getPath: () => os.tmpdir(),
  },
}));

vi.mock("electron-store", () => ({
  default: class MockStore {
    private readonly key: string;
    private data: Record<string, unknown>;

    constructor(options: { name?: string; cwd?: string; defaults?: Record<string, unknown> }) {
      this.key = `${options.cwd || "default"}::${options.name || "store"}`;
      const existing = testState.stores.get(this.key);
      this.data = existing ? { ...existing } : { ...(options.defaults || {}) };
      testState.stores.set(this.key, this.data);
    }

    get store(): Record<string, unknown> {
      return this.data;
    }

    set store(value: Record<string, unknown>) {
      this.data = { ...value };
      testState.stores.set(this.key, this.data);
    }

    get(key: string): unknown {
      return this.data[key];
    }

    set(key: string, value: unknown): void {
      this.data[key] = value;
      testState.stores.set(this.key, this.data);
    }

    clear(): void {
      this.data = {};
      testState.stores.set(this.key, this.data);
    }
  },
}));

vi.mock("../shared/logger", () => ({
  createLogger: () => ({
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  }),
  configureLogDirectory: vi.fn(),
}));

vi.mock("../agent/runtime/knowledgeRuntime", () => ({
  reloadKnowledgeRuntime: (...args: unknown[]) => testState.reloadKnowledge(...args),
  resetKnowledgeRuntime: (...args: unknown[]) => testState.resetKnowledge(...args),
}));

vi.mock("./settingsDataPath", async () => {
  const actual = await vi.importActual<typeof import("./settingsDataPath")>("./settingsDataPath");
  return {
    ...actual,
    getActiveDataPath: () => testState.activeDataPath,
    setConfiguredDataPath: (dataPath: string) => {
      testState.activeDataPath = path.resolve(dataPath);
    },
    logUserDataPathMigrateFailure: (...args: unknown[]) => testState.logMigrateFailure(...args),
  };
});

vi.mock("./localDataProtection/localDataLifecycle", () => ({
  afterDataPathMigrated: vi.fn(async () => ({ oldRootCleared: true })),
  bootstrapLocalDataProtection: vi.fn(async () => undefined),
}));

import {
  closeStateRuntimeStore,
  getSessionStoreInstance,
  migrateDataPath,
  setAgentLoopsGetter,
} from "./settingsManager";
import { normalizePathForCompare } from "./settingsDataPath";

describe("settingsManager.migrateDataPath post-commit runtime rebind failure", () => {
  const tempRoots: string[] = [];
  let previousSessionStore: ReturnType<typeof getSessionStoreInstance>;
  let sessionUpdates: unknown[] = [];
  let stateUpdates: unknown[] = [];

  beforeEach(() => {
    testState.stores.clear();
    testState.reloadKnowledge.mockReset();
    testState.resetKnowledge.mockReset();
    testState.logMigrateFailure.mockReset();
    sessionUpdates = [];
    stateUpdates = [];

    const root = fs.mkdtempSync(path.join(os.tmpdir(), "migrate-data-path-"));
    tempRoots.push(root);
    const current = path.join(root, "current");
    seedDataRoot(current);
    testState.activeDataPath = current;

    previousSessionStore = getSessionStoreInstance();
    const markerPath = path.join(current, "sessions", "marker.jsonl");
    previousSessionStore.resumeWrites();
    fs.appendFileSync(markerPath, "before-migration\n", "utf8");

    setAgentLoopsGetter(() => [
      {
        getIsRunning: () => false,
        updateSessionStore: (store: unknown) => {
          sessionUpdates.push(store);
        },
        updateStateRuntimeStore: (store: unknown) => {
          stateUpdates.push(store);
        },
      } as never,
    ]);
  });

  afterEach(async () => {
    setAgentLoopsGetter(() => []);
    await closeStateRuntimeStore().catch(() => {});
    testState.resetKnowledge();
    for (const root of tempRoots.splice(0)) {
      await fs.promises.rm(root, { recursive: true, force: true, maxRetries: 10, retryDelay: 50 });
    }
  });

  it("rolls back after commit when knowledge runtime rebind fails", async () => {
    const current = testState.activeDataPath;
    const next = path.join(path.dirname(current), "next");
    const oldOfficeBackup = path.join(current, "office-backups", "book.xlsx");
    const oldWorkflow = path.join(current, "office-automation", "workflows", "flow.json");
    const oldTransaction = path.join(current, "office-automation", "transactions", "tx.json");
    const oldMarker = path.join(current, "sessions", "marker.jsonl");

    testState.reloadKnowledge.mockImplementation(async (_config, dataRoot: string) => {
      if (normalizePathForCompare(dataRoot) === normalizePathForCompare(next)) {
        return { store: null, error: "forced knowledge rebind failure" };
      }
      return {
        store: { isInitialized: () => true, root: dataRoot },
        error: undefined,
      };
    });

    const result = await migrateDataPath(next);

    expect(result).toEqual({
      success: false,
      error: "forced knowledge rebind failure",
    });
    expect(testState.logMigrateFailure).toHaveBeenCalled();
    expect(normalizePathForCompare(testState.activeDataPath)).toBe(
      normalizePathForCompare(current),
    );

    // Old root Office trees remain intact; target cleaned (no split).
    expect(fs.readFileSync(oldMarker, "utf8")).toContain("before-migration");
    expect(fs.readFileSync(oldOfficeBackup, "utf8")).toBe("backup");
    expect(fs.readFileSync(oldWorkflow, "utf8")).toBe("workflow");
    expect(fs.readFileSync(oldTransaction, "utf8")).toBe("transaction");
    expect(fs.existsSync(next)).toBe(false);
    expect(fs.existsSync(path.join(next, "office-backups", "book.xlsx"))).toBe(false);
    expect(fs.existsSync(path.join(next, "office-automation", "workflows", "flow.json"))).toBe(
      false,
    );

    // Previous SessionStore is restored; production write must succeed (resumeWrites).
    expect(sessionUpdates[sessionUpdates.length - 1]).toBe(previousSessionStore);
    expect(getSessionStoreInstance()).toBe(previousSessionStore);
    const restoredThread = await previousSessionStore.createThread(
      "openai",
      "migrate-rollback-model",
    );
    await previousSessionStore.flushRolloutWrites();
    const rolloutPath = previousSessionStore.getRolloutPath(restoredThread.metadata.threadId);
    expect(normalizePathForCompare(rolloutPath).startsWith(normalizePathForCompare(current))).toBe(
      true,
    );
    expect(fs.existsSync(rolloutPath)).toBe(true);
    const loaded = await previousSessionStore.loadThread(restoredThread.metadata.threadId);
    expect(loaded?.metadata.threadId).toBe(restoredThread.metadata.threadId);
    expect(loaded?.metadata.model).toBe("migrate-rollback-model");

    // StateRuntime rebound to old root paths.
    const restoredState = stateUpdates[stateUpdates.length - 1] as {
      getDatabasePaths: () => {
        state: string;
        logs: string;
        goals: string;
        memories: string;
      };
    };
    expect(typeof restoredState.getDatabasePaths).toBe("function");
    const dbPaths = restoredState.getDatabasePaths();
    const expectedRuntimeRoot = path.join(current, "sessions", "state-runtime");
    expect(normalizePathForCompare(dbPaths.state)).toBe(
      normalizePathForCompare(path.join(expectedRuntimeRoot, "state.db")),
    );
    expect(normalizePathForCompare(dbPaths.logs)).toBe(
      normalizePathForCompare(path.join(expectedRuntimeRoot, "logs.db")),
    );
    expect(normalizePathForCompare(dbPaths.goals)).toBe(
      normalizePathForCompare(path.join(expectedRuntimeRoot, "goals.db")),
    );
    expect(normalizePathForCompare(dbPaths.memories)).toBe(
      normalizePathForCompare(path.join(expectedRuntimeRoot, "memories.db")),
    );

    // Knowledge restore rebind targets the old root.
    const reloadRoots = testState.reloadKnowledge.mock.calls.map((call) =>
      normalizePathForCompare(String(call[1])),
    );
    expect(reloadRoots).toContain(normalizePathForCompare(next));
    expect(reloadRoots[reloadRoots.length - 1]).toBe(normalizePathForCompare(current));
  });
});

function seedDataRoot(root: string): void {
  writeFile(root, "settings/excel-ai-settings.json", "{}");
  writeFile(root, "sessions/marker.jsonl", "");
  writeFile(root, "sessions/state-runtime/.keep", "");
  writeFile(root, "knowledge/.keep", "");
  writeFile(root, "office-backups/book.xlsx", "backup");
  writeFile(root, "office-automation/workflows/flow.json", "workflow");
  writeFile(root, "office-automation/transactions/tx.json", "transaction");
  writeFile(root, "logs/app.log", "log");
}

function writeFile(root: string, relativePath: string, content: string): void {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content, "utf8");
}
