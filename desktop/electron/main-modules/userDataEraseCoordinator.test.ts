import { describe, expect, it, vi } from "vitest";

import { USER_DATA_ERASE_CONFIRMATION } from "./userDataErase";
import { runUserDataErase, type UserDataEraseCoordinatorDeps } from "./userDataEraseCoordinator";

describe("user data erase coordinator", () => {
  it("requires exact confirmation and refuses active data operations or agents", async () => {
    const invalid = deps();
    await expect(runUserDataErase("erase", invalid)).resolves.toMatchObject({ success: false });
    expect(invalid.getSessionStore).not.toHaveBeenCalled();

    const busy = deps({ isBusy: () => true });
    await expect(runUserDataErase(USER_DATA_ERASE_CONFIRMATION, busy)).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("正在进行中"),
    });

    const running = deps({ hasRunningAgent: () => true });
    await expect(runUserDataErase(USER_DATA_ERASE_CONFIRMATION, running)).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("当前会话"),
    });
  });

  it("quiesces runtimes, clears settings, erases managed data, and restores in order", async () => {
    const calls: string[] = [];
    const coordinator = deps({
      setBusy: (value) => {
        calls.push(`busy:${value}`);
      },
      getSessionStore: vi.fn(() => ({
        suspendWrites: () => {
          calls.push("suspend");
        },
        resumeWrites: () => {
          calls.push("resume");
        },
        flushRolloutWrites: async () => {
          calls.push("flush");
        },
      })),
      resetAgents: async () => {
        calls.push("reset-agents");
      },
      closeStateRuntime: async () => {
        calls.push("close-state");
      },
      resetKnowledgeRuntime: () => {
        calls.push("reset-knowledge");
      },
      clearSettings: () => {
        calls.push("clear-settings");
      },
      eraseManagedData: async () => {
        calls.push("erase");
        return { erasedCategories: ["sessions"], errors: [] };
      },
      restoreRuntimes: async () => {
        calls.push("restore");
      },
    });

    await expect(runUserDataErase(USER_DATA_ERASE_CONFIRMATION, coordinator)).resolves.toEqual({
      success: true,
      erasedCategories: ["settings", "sessions"],
      errors: [],
    });
    expect(calls).toEqual([
      "busy:true",
      "suspend",
      "flush",
      "reset-agents",
      "close-state",
      "reset-knowledge",
      "clear-settings",
      "erase",
      "restore",
      "resume",
      "busy:false",
    ]);
  });

  it("reports partial deletion and runtime recovery failures without claiming success", async () => {
    const coordinator = deps({
      eraseManagedData: async () => ({
        erasedCategories: ["knowledge"],
        errors: ["sessions: access denied"],
      }),
      restoreRuntimes: async () => {
        throw new Error("runtime unavailable");
      },
    });

    await expect(
      runUserDataErase(USER_DATA_ERASE_CONFIRMATION, coordinator),
    ).resolves.toMatchObject({
      success: false,
      erasedCategories: ["settings", "knowledge"],
      error: expect.stringContaining("runtime unavailable"),
    });
  });

  it("reports cleared settings when managed directory erasure throws", async () => {
    const coordinator = deps({
      eraseManagedData: async () => {
        throw new Error("data root unavailable");
      },
    });

    await expect(
      runUserDataErase(USER_DATA_ERASE_CONFIRMATION, coordinator),
    ).resolves.toMatchObject({
      success: false,
      erasedCategories: ["settings"],
      error: expect.stringContaining("data root unavailable"),
    });
  });
});

function deps(
  overrides: Partial<UserDataEraseCoordinatorDeps> = {},
): UserDataEraseCoordinatorDeps & { getSessionStore: ReturnType<typeof vi.fn> } {
  const getSessionStore = vi.fn(() => ({
    suspendWrites: vi.fn(),
    resumeWrites: vi.fn(),
    flushRolloutWrites: vi.fn(async () => undefined),
  }));
  return {
    isBusy: () => false,
    setBusy: vi.fn(),
    hasRunningAgent: () => false,
    getDataPath: () => "C:\\data",
    getSessionStore,
    resetAgents: vi.fn(async () => undefined),
    closeStateRuntime: vi.fn(async () => undefined),
    resetKnowledgeRuntime: vi.fn(),
    clearSettings: vi.fn(),
    restoreRuntimes: vi.fn(async () => undefined),
    eraseManagedData: vi.fn(async () => ({ erasedCategories: [], errors: [] })),
    ...overrides,
  } as UserDataEraseCoordinatorDeps & { getSessionStore: ReturnType<typeof vi.fn> };
}
