import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import { eraseManagedUserData } from "./userDataErase";

describe("managed user data erasure", () => {
  const roots: string[] = [];

  afterEach(() => {
    for (const root of roots.splice(0)) fs.rmSync(root, { recursive: true, force: true });
  });

  it("removes every managed content directory without deleting settings or unknown files", async () => {
    const root = makeRoot();
    for (const relativePath of [
      "sessions/thread.jsonl",
      "knowledge/knowledge.db",
      "office-backups/backup.xlsx",
      "office-automation/workflows/flow.json",
      "logs/app.log",
      "temp/ocr.txt",
      "settings/excel-ai-settings.json",
      "operator-notes.txt",
    ])
      writeFile(root, relativePath);

    const report = await eraseManagedUserData(root);

    expect(report.errors).toEqual([]);
    expect(report.erasedCategories).toHaveLength(6);
    expect(fs.existsSync(path.join(root, "sessions"))).toBe(false);
    expect(fs.existsSync(path.join(root, "settings/excel-ai-settings.json"))).toBe(true);
    expect(fs.existsSync(path.join(root, "operator-notes.txt"))).toBe(true);
  });

  it("continues other categories and reports a failed deletion", async () => {
    const root = makeRoot();
    writeFile(root, "sessions/thread.jsonl");
    writeFile(root, "knowledge/knowledge.db");

    const report = await eraseManagedUserData(root, {
      removeDirectory: async (directory) => {
        if (path.basename(directory) === "sessions") throw new Error("access denied");
        await fs.promises.rm(directory, { recursive: true });
      },
    });

    expect(report.errors).toEqual(["sessions: access denied"]);
    expect(fs.existsSync(path.join(root, "sessions/thread.jsonl"))).toBe(true);
    expect(fs.existsSync(path.join(root, "knowledge"))).toBe(false);
  });

  it("refuses managed directory links without deleting their targets", async () => {
    const outside = makeRoot();
    const root = makeRoot();
    writeFile(outside, "keep.txt");
    fs.symlinkSync(outside, path.join(root, "sessions"), "junction");

    const report = await eraseManagedUserData(root);

    expect(report.errors).toEqual(["sessions: 拒绝删除符号链接或联接"]);
    expect(fs.existsSync(path.join(outside, "keep.txt"))).toBe(true);
    expect(fs.existsSync(path.join(root, "sessions"))).toBe(true);
  });

  it("refuses a filesystem root before evaluating managed children", async () => {
    const root = path.parse(makeRoot()).root;
    await expect(eraseManagedUserData(root)).rejects.toThrow("拒绝擦除磁盘根目录");
  });

  function makeRoot(): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "user-data-erase-"));
    roots.push(root);
    return root;
  }
});

function writeFile(root: string, relativePath: string): void {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, "data");
}
