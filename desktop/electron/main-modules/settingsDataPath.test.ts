import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

vi.mock("electron", () => ({
  app: {
    isPackaged: false,
    getPath: () => os.tmpdir(),
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

import { migrateLegacyDataDirectorySync } from "./settingsDataPath";

const temporaryDirectories: string[] = [];

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
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
