import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import {
  cleanupPreparedDataPathMigration,
  commitPreparedDataPathMigration,
  prepareDataPathMigration,
} from "./dataPathMigration";

describe("transactional data-path staging", () => {
  const tempRoots: string[] = [];

  afterEach(() => {
    for (const root of tempRoots.splice(0)) {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });

  it("copies settings, sessions, knowledge, Office state, backups, logs and temp data", async () => {
    const root = makeRoot();
    const current = path.join(root, "current");
    const target = path.join(root, "next");
    const files = [
      "settings/excel-ai-settings.json",
      "sessions/state-runtime/state.db",
      "knowledge/knowledge.db",
      "office-backups/book.xlsx",
      "office-automation/workflows/flow.json",
      "office-automation/transactions/tx.json",
      "logs/app.log",
      "temp/ocr/input.txt",
    ];
    for (const relativePath of files) writeFile(current, relativePath, relativePath);

    const prepared = await prepareDataPathMigration(current, target);
    await commitPreparedDataPathMigration(prepared);

    for (const relativePath of files) {
      expect(fs.readFileSync(path.join(target, relativePath), "utf8")).toBe(relativePath);
    }
  });

  it("refuses to merge into a non-empty target", async () => {
    const root = makeRoot();
    const current = path.join(root, "current");
    const target = path.join(root, "next");
    writeFile(current, "sessions/thread.jsonl", "source");
    writeFile(target, "unrelated.txt", "keep");

    await expect(prepareDataPathMigration(current, target)).rejects.toThrow("必须为空");
    expect(fs.readFileSync(path.join(target, "unrelated.txt"), "utf8")).toBe("keep");
  });

  it("rejects UNC network targets by default", async () => {
    const root = makeRoot();
    const current = path.join(root, "current");
    writeFile(current, "sessions/thread.jsonl", "source");

    await expect(prepareDataPathMigration(current, "\\\\server\\share\\wenge-data"))
      .rejects.toThrow("UNC");
  });

  it("removes staging data and restores an originally empty target on rollback", async () => {
    const root = makeRoot();
    const current = path.join(root, "current");
    const target = path.join(root, "next");
    writeFile(current, "sessions/thread.jsonl", "source");
    fs.mkdirSync(target, { recursive: true });

    const prepared = await prepareDataPathMigration(current, target);
    await commitPreparedDataPathMigration(prepared);
    await cleanupPreparedDataPathMigration(prepared, true);

    expect(fs.existsSync(target)).toBe(true);
    expect(fs.readdirSync(target)).toEqual([]);
  });

  function makeRoot(): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "data-migration-"));
    tempRoots.push(root);
    return root;
  }
});

function writeFile(root: string, relativePath: string, content: string): void {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content);
}
