import { access, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import type { OfficeActionBridge } from "../contracts/office";
import type { OfficeActionResult } from "./types";
import {
  beginOfficeTransaction,
  finalizeOfficeTransaction,
  getOfficeTransaction,
  recordOfficeTransactionResult,
  redoOfficeTransaction,
  saveOfficeTransaction,
  undoOfficeTransaction,
} from "./transactionJournal";

describe("Office transaction journal", () => {
  const temporaryRoots: string[] = [];

  afterEach(async () => {
    await Promise.all(
      temporaryRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })),
    );
  });

  it("restores existing files and deletes declared new outputs as one transaction", async () => {
    const root = await createRoot(temporaryRoots);
    const sourcePath = path.join(root, "book.xlsx");
    const outputPath = path.join(root, "report.docx");
    const journalRoot = path.join(root, "journal");
    await writeFile(sourcePath, "before", "utf8");
    const step = {
      app: "excel" as const,
      action: "edit" as const,
      operation: "buildReport",
      filePath: sourcePath,
      outputPath,
    };
    const record = await beginOfficeTransaction({ root: journalRoot, steps: [step] });

    await writeFile(sourcePath, "after", "utf8");
    await writeFile(outputPath, "new report", "utf8");
    await recordOfficeTransactionResult(journalRoot, record, actionResult(step, outputPath));
    await finalizeOfficeTransaction(journalRoot, record);

    const undone = await undoOfficeTransaction(journalRoot, record.id);

    expect(undone.status).toBe("undone");
    expect(await readFile(sourcePath, "utf8")).toBe("before");
    await expect(access(outputPath)).rejects.toThrow();
    expect(undone.artifacts).toEqual([outputPath]);
    expect(undone.changes).toEqual([
      { kind: "create", target: outputPath, detail: "生成 Word 报告" },
    ]);
  });

  it("restores the deterministic after snapshot without re-executing tools", async () => {
    const root = await createRoot(temporaryRoots);
    const sourcePath = path.join(root, "book.xlsx");
    const outputPath = path.join(root, "report.pptx");
    const journalRoot = path.join(root, "journal");
    await writeFile(sourcePath, "before", "utf8");
    const step = {
      app: "presentation" as const,
      action: "insert" as const,
      operation: "buildSlides",
      filePath: sourcePath,
      outputPath,
    };
    const record = await beginOfficeTransaction({ root: journalRoot, steps: [step] });
    await writeFile(outputPath, "first", "utf8");
    await recordOfficeTransactionResult(journalRoot, record, actionResult(step, outputPath));
    await finalizeOfficeTransaction(journalRoot, record);
    await undoOfficeTransaction(journalRoot, record.id);
    const executeAction = vi.fn(async () => {
      await writeFile(outputPath, "should-not-run", "utf8");
      return actionResult(step, outputPath);
    });
    const bridge: OfficeActionBridge = { executeAction };

    const redone = await redoOfficeTransaction(journalRoot, record.id, bridge);

    expect(redone.status).toBe("applied");
    expect(redone.results).toHaveLength(1);
    expect(redone.changes).toHaveLength(1);
    expect(await readFile(outputPath, "utf8")).toBe("first");
    expect(executeAction).not.toHaveBeenCalled();
  });

  it("delegates open-document restore to a coordinator after flushing current files", async () => {
    const root = await createRoot(temporaryRoots);
    const sourcePath = path.join(root, "book.xlsx");
    const journalRoot = path.join(root, "journal");
    await writeFile(sourcePath, "before", "utf8");
    const step = {
      app: "excel" as const,
      action: "edit" as const,
      operation: "format",
      filePath: sourcePath,
    };
    const record = await beginOfficeTransaction({ root: journalRoot, steps: [step] });
    await writeFile(sourcePath, "after", "utf8");
    await recordOfficeTransactionResult(journalRoot, record, {
      status: "done",
      engine: "com",
      ...step,
      summary: "done",
      changes: [],
    });
    await finalizeOfficeTransaction(journalRoot, record);
    const order: string[] = [];
    const prepareFiles = vi.fn(async () => {
      order.push("prepare");
    });
    const restoreFiles = vi.fn(
      async (files: Array<{ filePath: string; existed: boolean; snapshotPath?: string }>) => {
        order.push("restore");
        for (const file of files) {
          if (file.existed)
            await writeFile(file.filePath, await readFile(file.snapshotPath!, "utf8"), "utf8");
          else await rm(file.filePath, { force: true });
        }
      },
    );

    await undoOfficeTransaction(journalRoot, record.id, { prepareFiles, restoreFiles });
    expect(await readFile(sourcePath, "utf8")).toBe("before");
    await redoOfficeTransaction(journalRoot, record.id, undefined, { prepareFiles, restoreFiles });

    expect(await readFile(sourcePath, "utf8")).toBe("after");
    expect(order).toEqual(["prepare", "restore", "prepare", "restore"]);
    expect(restoreFiles.mock.calls[0][0][0].snapshotPath).toContain(record.id);
  });

  it("rejects completion when an output was not declared before execution", async () => {
    const root = await createRoot(temporaryRoots);
    const sourcePath = path.join(root, "book.xlsx");
    const undeclaredPath = path.join(root, "existing.docx");
    const journalRoot = path.join(root, "journal");
    await writeFile(sourcePath, "before", "utf8");
    await writeFile(undeclaredPath, "existing", "utf8");
    const step = {
      app: "excel" as const,
      action: "edit" as const,
      operation: "unknownOutput",
      filePath: sourcePath,
    };
    const record = await beginOfficeTransaction({ root: journalRoot, steps: [step] });
    await recordOfficeTransactionResult(journalRoot, record, actionResult(step, undeclaredPath));

    await expect(finalizeOfficeTransaction(journalRoot, record)).rejects.toThrow("未声明产物");

    expect(await readFile(undeclaredPath, "utf8")).toBe("existing");
    expect(record.artifacts).toEqual([undeclaredPath]);
    expect(record.snapshots.some((snapshot) => snapshot.filePath === undeclaredPath)).toBe(false);
  });

  it("blocks undo after external edits and allows an explicit forced restore", async () => {
    const root = await createRoot(temporaryRoots);
    const sourcePath = path.join(root, "book.xlsx");
    const journalRoot = path.join(root, "journal");
    await writeFile(sourcePath, "before", "utf8");
    const step = {
      app: "excel" as const,
      action: "edit" as const,
      operation: "format",
      filePath: sourcePath,
    };
    const record = await beginOfficeTransaction({ root: journalRoot, steps: [step] });
    await writeFile(sourcePath, "after", "utf8");
    await recordOfficeTransactionResult(journalRoot, record, {
      status: "done",
      engine: "com",
      ...step,
      summary: "done",
      changes: [{ kind: "style", target: "range:Sheet1!A1", detail: "格式化" }],
    });
    await finalizeOfficeTransaction(journalRoot, record);
    await writeFile(sourcePath, "external edit", "utf8");

    const conflicted = await undoOfficeTransaction(journalRoot, record.id);
    expect(conflicted).toMatchObject({ status: "conflicted", conflictBaseStatus: "applied" });
    expect(conflicted.conflicts?.[0].reason).toContain("事务外修改");
    expect(await readFile(sourcePath, "utf8")).toBe("external edit");

    const forced = await undoOfficeTransaction(journalRoot, record.id, { force: true });
    expect(forced.status).toBe("undone");
    expect(await readFile(sourcePath, "utf8")).toBe("before");
  });

  it("rejects malformed transaction IDs before resolving paths", async () => {
    const root = await createRoot(temporaryRoots);
    await expect(getOfficeTransaction(root, "../../outside")).rejects.toThrow("事务 ID 无效");
    await expect(
      getOfficeTransaction(root, "------------------------------------"),
    ).rejects.toThrow("事务 ID 无效");
  });
});

async function createRoot(roots: string[]): Promise<string> {
  const root = await mkdtemp(path.join(os.tmpdir(), "office-journal-"));
  roots.push(root);
  return root;
}

function actionResult(
  step: {
    app: "excel" | "word" | "presentation";
    action: "edit" | "insert";
    operation: string;
    filePath: string;
    outputPath?: string;
  },
  outputPath: string,
): OfficeActionResult {
  return {
    status: "done",
    engine: "com",
    ...step,
    outputPath,
    summary: "done",
    changes: [
      {
        kind: "create",
        target: outputPath,
        detail: step.app === "presentation" ? "生成演示文稿" : "生成 Word 报告",
      },
    ],
  };
}
