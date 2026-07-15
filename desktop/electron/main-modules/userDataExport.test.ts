import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, describe, expect, it } from "vitest";

import { exportUserDataDirectory } from "./userDataExport";

describe("user data privacy export", () => {
  const tempRoots: string[] = [];

  afterEach(() => {
    for (const root of tempRoots.splice(0)) {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });

  it("exports all managed categories while replacing encrypted settings with a safe snapshot", async () => {
    const root = makeRoot();
    const source = path.join(root, "source");
    const target = path.join(root, "export");
    writeFile(source, "settings/excel-ai-settings.json", "encrypted-api-key");
    writeFile(source, "sessions/2026/07/15/thread.jsonl", "conversation");
    writeFile(source, "sessions/state-runtime/memories.db", "memory");
    writeFile(source, "knowledge/knowledge.db", "knowledge");
    writeFile(source, "office-backups/backup.xlsx", "backup");
    writeFile(source, "office-automation/workflows/flow.json", "workflow");
    writeFile(source, "logs/app.log", "log");
    writeFile(source, "temp/ocr/input.txt", "temporary");

    const result = await exportUserDataDirectory({
      sourceDataPath: source,
      targetPath: target,
      sanitizedSettings: {
        theme: "dark",
        apiKey: "••••••••",
        compactionConfig: { remoteCompactApiKey: "plaintext-compact-secret" },
      },
      now: new Date("2026-07-15T04:00:00.000Z"),
    });

    expect(result.exportPath).toBe(path.resolve(target));
    expect(result.categories).toContain("sessions");
    expect(fs.readFileSync(path.join(target, "sessions/2026/07/15/thread.jsonl"), "utf8")).toBe(
      "conversation",
    );
    expect(fs.existsSync(path.join(target, "settings/excel-ai-settings.json"))).toBe(false);
    const settings = readJson(path.join(target, "settings/privacy-export-settings.json"));
    expect(settings).toMatchObject({
      exportedAt: "2026-07-15T04:00:00.000Z",
      credentials: "omitted",
      settings: {
        theme: "dark",
        apiKey: "[REDACTED]",
        compactionConfig: { remoteCompactApiKey: "[REDACTED]" },
      },
    });
    expect(JSON.stringify(settings)).not.toContain("encrypted-api-key");
    expect(JSON.stringify(settings)).not.toContain("plaintext-compact-secret");
    expect(readJson(path.join(target, "privacy-export-manifest.json"))).toMatchObject({
      schemaVersion: 1,
      categories: expect.arrayContaining(["knowledge-index", "office-backups-and-automation"]),
    });
  });

  it("does not merge into a non-empty destination", async () => {
    const root = makeRoot();
    const source = path.join(root, "source");
    const target = path.join(root, "export");
    writeFile(source, "sessions/thread.jsonl", "conversation");
    writeFile(target, "unrelated.txt", "keep");

    await expect(
      exportUserDataDirectory({
        sourceDataPath: source,
        targetPath: target,
        sanitizedSettings: {},
      }),
    ).rejects.toThrow("必须为空");
    expect(fs.readFileSync(path.join(target, "unrelated.txt"), "utf8")).toBe("keep");
  });

  function makeRoot(): string {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "user-data-export-"));
    tempRoots.push(root);
    return root;
  }
});

function writeFile(root: string, relativePath: string, content: string): void {
  const filePath = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, content);
}

function readJson(filePath: string): Record<string, unknown> {
  return JSON.parse(fs.readFileSync(filePath, "utf8")) as Record<string, unknown>;
}
