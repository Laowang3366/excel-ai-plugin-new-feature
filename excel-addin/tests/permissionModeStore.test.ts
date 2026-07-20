import { describe, expect, it, beforeEach } from "vitest";
import {
  DEFAULT_PERMISSION_MODE,
  PermissionModeStore,
  loadPermissionMode,
  persistPermissionMode,
  PERMISSION_MODE_PERSISTENCE_KEY,
  resetBrowserPermissionModeStoreForTests,
  getBrowserPermissionModeStore,
} from "../shared/agentChat";

class MemoryStorage {
  private map = new Map<string, string>();
  getItem(key: string): string | null {
    return this.map.has(key) ? this.map.get(key)! : null;
  }
  setItem(key: string, value: string): void {
    this.map.set(key, value);
  }
  clear(): void {
    this.map.clear();
  }
}

describe("PermissionModeStore persistence", () => {
  beforeEach(() => {
    resetBrowserPermissionModeStoreForTests();
  });

  it("defaults to auto_approve_safe without storage", () => {
    const store = new PermissionModeStore();
    expect(store.get()).toBe(DEFAULT_PERMISSION_MODE);
  });

  it("persists and reloads valid modes", () => {
    const memory = new MemoryStorage();
    const a = new PermissionModeStore(memory);
    expect(a.set("normal")).toBe("normal");
    expect(a.get()).toBe("normal");
    const b = new PermissionModeStore(memory);
    expect(b.get()).toBe("normal");
    b.set("confirm_all");
    expect(new PermissionModeStore(memory).get()).toBe("confirm_all");
  });

  it("illegal / corrupt persistence recovers to default", () => {
    const memory = new MemoryStorage();
    memory.setItem(PERMISSION_MODE_PERSISTENCE_KEY, "{not-json");
    expect(loadPermissionMode(memory)).toBe(DEFAULT_PERMISSION_MODE);

    memory.setItem(
      PERMISSION_MODE_PERSISTENCE_KEY,
      JSON.stringify({ version: 1, permissionMode: "full_auto" }),
    );
    expect(loadPermissionMode(memory)).toBe(DEFAULT_PERMISSION_MODE);

    memory.setItem(PERMISSION_MODE_PERSISTENCE_KEY, "null");
    expect(loadPermissionMode(memory)).toBe(DEFAULT_PERMISSION_MODE);

    const store = new PermissionModeStore(memory);
    expect(store.get()).toBe(DEFAULT_PERMISSION_MODE);
    // set rejects illegal and stores default
    expect(store.set("totally-invalid")).toBe(DEFAULT_PERMISSION_MODE);
    expect(JSON.parse(memory.getItem(PERMISSION_MODE_PERSISTENCE_KEY)!)).toEqual(
      {
        version: 1,
        permissionMode: DEFAULT_PERMISSION_MODE,
      },
    );
  });

  it("does not use provider secret storage keys", () => {
    const memory = new MemoryStorage();
    persistPermissionMode(memory, "normal");
    expect(memory.getItem(PERMISSION_MODE_PERSISTENCE_KEY)).toBeTruthy();
    expect(memory.getItem("wengge.excel-addin.provider-state")).toBeNull();
    expect(memory.getItem(PERMISSION_MODE_PERSISTENCE_KEY)).not.toMatch(
      /apiKey|secret|sk-/i,
    );
  });

  it("browser singleton is resettable for tests", () => {
    const first = getBrowserPermissionModeStore();
    first.set("normal");
    resetBrowserPermissionModeStoreForTests();
    const second = getBrowserPermissionModeStore();
    expect(second).not.toBe(first);
  });
});
