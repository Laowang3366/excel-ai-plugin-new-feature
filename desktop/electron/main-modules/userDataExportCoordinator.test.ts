import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";

import { runUserDataExport, type UserDataExportCoordinatorDeps } from "./userDataExportCoordinator";

describe("user data export coordinator", () => {
  const tempRoots: string[] = [];

  afterEach(() => {
    for (const root of tempRoots.splice(0)) {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });

  it("refuses export while another data operation or agent turn is active", async () => {
    const busy = deps({ isBusy: () => true });
    await expect(runUserDataExport("C:\\export", busy)).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("正在进行中"),
    });
    expect(busy.getSessionStore).not.toHaveBeenCalled();

    const running = deps({ hasRunningAgent: () => true });
    await expect(runUserDataExport("C:\\export", running)).resolves.toMatchObject({
      success: false,
      error: expect.stringContaining("当前会话"),
    });
    expect(running.getSessionStore).not.toHaveBeenCalled();
  });

  it("quiesces writers, exports, restores runtimes, and releases the operation lock", async () => {
    const root = makeRoot();
    const source = path.join(root, "source");
    const target = path.join(root, "export");
    writeFile(source, "sessions/thread.jsonl", "conversation");
    const calls: string[] = [];
    let busy = false;
    const coordinator = deps({
      isBusy: () => busy,
      setBusy: (value) => {
        busy = value;
        calls.push(`busy:${value}`);
      },
      getDataPath: () => source,
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
      closeStateRuntime: async () => {
        calls.push("close-state");
      },
      resetKnowledgeRuntime: () => {
        calls.push("reset-knowledge");
      },
      restoreRuntimes: async () => {
        calls.push("restore");
      },
    });

    await expect(runUserDataExport(target, coordinator)).resolves.toMatchObject({
      success: true,
      exportPath: path.resolve(target),
    });
    expect(calls).toEqual([
      "busy:true",
      "suspend",
      "flush",
      "close-state",
      "reset-knowledge",
      "resume",
      "restore",
      "busy:false",
    ]);
    expect(busy).toBe(false);
  });

  it("reports runtime recovery failure without hiding a completed export", async () => {
    const root = makeRoot();
    const source = path.join(root, "source");
    const target = path.join(root, "export");
    writeFile(source, "sessions/thread.jsonl", "conversation");
    const coordinator = deps({
      getDataPath: () => source,
      restoreRuntimes: async () => {
        throw new Error("runtime unavailable");
      },
    });

    await expect(runUserDataExport(target, coordinator)).resolves.toMatchObject({
      success: false,
      exportPath: path.resolve(target),
      error: expect.stringContaining("runtime unavailable"),
    });
    expect(fs.existsSync(path.join(target, "privacy-export-manifest.json"))).toBe(true);
  });

  function makeRoot(): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "user-data-export-coordinator-"));
    tempRoots.push(root);
    return root;
  }
});

function deps(
  overrides: Partial<UserDataExportCoordinatorDeps> = {},
): UserDataExportCoordinatorDeps & { getSessionStore: ReturnType<typeof vi.fn> } {
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
    getSanitizedSettings: () => ({ theme: "dark" }),
    getSessionStore,
    closeStateRuntime: vi.fn(async () => undefined),
    resetKnowledgeRuntime: vi.fn(),
    restoreRuntimes: vi.fn(async () => undefined),
    ...overrides,
  } as UserDataExportCoordinatorDeps & { getSessionStore: ReturnType<typeof vi.fn> };
}

function writeFile(root: string, relativePath: string, content: string): void {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content);
}
