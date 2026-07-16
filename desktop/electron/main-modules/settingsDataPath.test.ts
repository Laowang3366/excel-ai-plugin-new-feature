import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

const temporaryDirectories: string[] = [];
const testState = vi.hoisted(() => ({
  userDataRoot: "",
  dataPathWarn: vi.fn(),
  dataPathError: vi.fn(),
}));
testState.userDataRoot = os.tmpdir();

vi.mock("electron", () => ({
  app: {
    isPackaged: false,
    getPath: (name: string) => {
      if (name === "userData") return testState.userDataRoot;
      return os.tmpdir();
    },
  },
}));

vi.mock("electron-store", () => ({
  default: class MockStore {
    get(): undefined {
      return undefined;
    }

    set(): void {}
  },
}));

vi.mock("../shared/logger", () => ({
  createLogger: () => ({
    info: vi.fn(),
    warn: testState.dataPathWarn,
    error: testState.dataPathError,
    debug: vi.fn(),
  }),
}));

import {
  getActiveDataPath,
  logUserDataPathMigrateFailure,
  migrateLegacyDataDirectorySync,
} from "./settingsDataPath";

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
  testState.dataPathWarn.mockClear();
  testState.dataPathError.mockClear();
  vi.restoreAllMocks();
});

function createTemporaryRoot(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "settings-data-path-"));
  temporaryDirectories.push(root);
  return root;
}

describe("migrateLegacyDataDirectorySync", () => {
  it("copies through staging and atomically replaces an empty target", () => {
    const root = createTemporaryRoot();
    const source = path.join(root, "legacy-data");
    const target = path.join(root, "user-data");
    fs.mkdirSync(path.join(source, "sessions"), { recursive: true });
    fs.writeFileSync(path.join(source, "sessions", "thread.jsonl"), "record", "utf8");
    fs.mkdirSync(target, { recursive: true });

    migrateLegacyDataDirectorySync(source, target);

    expect(fs.readFileSync(path.join(target, "sessions", "thread.jsonl"), "utf8")).toBe("record");
    expect(fs.readdirSync(root).some((entry) => entry.includes("legacy-migration"))).toBe(false);
  });

  it("refuses to overwrite a target that already contains data", () => {
    const root = createTemporaryRoot();
    const source = path.join(root, "legacy-data");
    const target = path.join(root, "user-data");
    fs.mkdirSync(source, { recursive: true });
    fs.mkdirSync(target, { recursive: true });
    fs.writeFileSync(path.join(source, "source.txt"), "source", "utf8");
    fs.writeFileSync(path.join(target, "existing.txt"), "existing", "utf8");

    expect(() => migrateLegacyDataDirectorySync(source, target)).toThrow("拒绝覆盖");
    expect(fs.readFileSync(path.join(target, "existing.txt"), "utf8")).toBe("existing");
  });
});

describe("logUserDataPathMigrateFailure", () => {
  it("logs a stable event for user-initiated migrate failures", () => {
    logUserDataPathMigrateFailure(new Error("stage failed"));
    expect(testState.dataPathError).toHaveBeenCalledWith(
      "用户主动数据目录迁移失败",
      expect.objectContaining({
        event: "desktop.data_path.user_migrate_failed",
        message: "stage failed",
      }),
    );
  });
});

describe("getActiveDataPath legacy auto-migrate events", () => {
  it("logs a stable event when automatic legacy migration fails", () => {
    const root = createTemporaryRoot();
    const installRoot = path.join(root, "install");
    testState.userDataRoot = path.join(root, "userData");
    const installData = path.join(installRoot, "data");
    const sessions = path.join(installData, "sessions");
    fs.mkdirSync(sessions, { recursive: true });
    fs.writeFileSync(path.join(sessions, "thread.jsonl"), "record", "utf8");
    // Junction forces the legacy migration guard to fail without mocking fs.cpSync.
    fs.symlinkSync(sessions, path.join(installData, "linked-sessions"), "junction");
    fs.mkdirSync(testState.userDataRoot, { recursive: true });

    const cwdSpy = vi.spyOn(process, "cwd").mockReturnValue(installRoot);
    try {
      getActiveDataPath();
      expect(testState.dataPathWarn).toHaveBeenCalledWith(
        "旧安装目录数据自动迁移失败",
        expect.objectContaining({
          event: "desktop.data_path.legacy_auto_migrate_failed",
          error: "旧数据目录包含符号链接或联接",
        }),
      );
    } finally {
      cwdSpy.mockRestore();
    }
  });
});
