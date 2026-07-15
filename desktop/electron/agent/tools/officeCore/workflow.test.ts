import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import type { OfficeActionBridge } from "../contracts/office";
import {
  listOfficeWorkflows,
  requestOfficeWorkflowCancellation,
  runOfficeWorkflow,
} from "./workflow";

describe("runOfficeWorkflow", () => {
  const temporaryRoots: string[] = [];

  afterEach(async () => {
    await Promise.all(temporaryRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
  });

  it("rejects advanced Excel steps without explicit semantic boundaries before execution", async () => {
    const bridge: OfficeActionBridge = { executeAction: vi.fn() };

    const result = await runOfficeWorkflow(bridge, [{
      app: "excel",
      action: "insert",
      operation: "createPivotTable",
      filePath: "C:/book.xlsx",
      target: "range:Sheet1!A1:B10",
      params: { rowFields: ["Department"] },
    }]);

    expect(result).toMatchObject({ status: "failed" });
    expect(result.error).toContain("interactive-pivot");
    expect(bridge.executeAction).not.toHaveBeenCalled();
  });

  it("rolls back in-place changes in reverse order when a later step fails", async () => {
    const executeAction = vi.fn(async (input) => {
      if (input.operation === "restoreBackup") {
        return { status: "done" as const, engine: "openxml" as const, ...input, summary: "restored", changes: [] };
      }
      const index = executeAction.mock.calls.length;
      return {
        status: index === 3 ? "failed" as const : "done" as const,
        engine: "com" as const,
        ...input,
        summary: index === 3 ? "failed" : "done",
        changes: [],
        data: { transaction: { sourcePath: input.filePath, backupPath: `C:/backups/${index}.bak` } },
      };
    });
    const bridge: OfficeActionBridge = { executeAction };

    const result = await runOfficeWorkflow(bridge, [
      { app: "excel", action: "style", operation: "applyWorkbookTemplate", filePath: "C:/book.xlsx" },
      { app: "word", action: "style", operation: "formatLongDocument", filePath: "C:/report.docx" },
      { app: "presentation", action: "style", operation: "layoutElements", filePath: "C:/slides.pptx" },
    ]);

    expect(result).toMatchObject({ status: "failed", completedSteps: 2, failedStep: 3 });
    expect(result.rollback.map((item) => item.step)).toEqual([3, 2, 1]);
    expect(executeAction.mock.calls.slice(3).map(([input]) => input.operation)).toEqual([
      "restoreBackup",
      "restoreBackup",
      "restoreBackup",
    ]);
  });

  it("allows intermediate outputs and records their artifacts", async () => {
    const bridge: OfficeActionBridge = { executeAction: vi.fn(async (input) => ({
      status: "done" as const,
      engine: "com" as const,
      ...input,
      outputPath: input.outputPath,
      summary: "done",
      changes: [],
    })) };
    const result = await runOfficeWorkflow(bridge, [
      { app: "excel", action: "edit", operation: "exportPdf", filePath: "C:/book.xlsx", outputPath: "C:/out/book.pdf" },
      { app: "word", action: "style", operation: "formatLongDocument", filePath: "C:/report.docx" },
    ]);

    expect(result).toMatchObject({ status: "done", completedSteps: 2 });
    expect(result.stepRecords[0].artifacts).toEqual(["C:/out/book.pdf"]);
    expect(bridge.executeAction).toHaveBeenCalledTimes(2);
  });

  it("pauses after an exception and resumes from the failed step", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-"));
    temporaryRoots.push(root);
    const workflowRoot = path.join(root, "workflows");
    const transactionRoot = path.join(root, "transactions");
    const sourcePath = path.join(root, "book.xlsx");
    const outputPath = path.join(root, "report.docx");
    await writeFile(sourcePath, "source", "utf8");
    let failSecondStep = true;
    const executeAction = vi.fn(async (input) => {
      if (input.operation === "buildWordReport" && failSecondStep) throw new Error("Word 暂时不可用");
      if (input.operation === "buildWordReport") await writeFile(outputPath, "report", "utf8");
      return {
        status: "done" as const,
        engine: "com" as const,
        ...input,
        outputPath: input.outputPath,
        summary: "done",
        changes: [],
      };
    });
    const bridge: OfficeActionBridge = { executeAction };
    const steps = [
      { app: "excel" as const, action: "edit" as const, operation: "configurePrint", filePath: sourcePath },
      { app: "word" as const, action: "insert" as const, operation: "buildWordReport", filePath: sourcePath, outputPath },
      { app: "presentation" as const, action: "insert" as const, operation: "buildSlides", filePath: sourcePath },
    ];

    const paused = await runOfficeWorkflow(bridge, steps, { workflowRoot, transactionRoot });

    expect(paused).toMatchObject({ status: "paused", completedSteps: 1, failedStep: 2, nextStep: 2 });
    expect(paused.error).toBe("Word 暂时不可用");
    expect(paused.workflowId).toBeTruthy();
    expect(paused.transactionId).toBeTruthy();
    failSecondStep = false;

    const resumed = await runOfficeWorkflow(bridge, [], {
      workflowRoot,
      transactionRoot,
      workflowId: paused.workflowId,
      resume: true,
    });

    expect(resumed).toMatchObject({ status: "done", completedSteps: 3 });
    expect(executeAction.mock.calls.map(([input]) => input.operation)).toEqual([
      "configurePrint",
      "buildWordReport",
      "buildWordReport",
      "buildSlides",
    ]);
    expect(resumed.stepRecords[1].artifacts).toEqual([outputPath]);
    expect(await readFile(outputPath, "utf8")).toBe("report");
  });

  it("resolves outputs, retries failures, runs safe parallel groups, and skips false conditions", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-advanced-"));
    temporaryRoots.push(root);
    const sourcePath = path.join(root, "source.xlsx");
    const preparedPath = path.join(root, "prepared.xlsx");
    const wordPath = path.join(root, "report.docx");
    const pptPath = path.join(root, "slides.pptx");
    await writeFile(sourcePath, "source", "utf8");
    let active = 0;
    let maxActive = 0;
    let wordAttempts = 0;
    const executeAction = vi.fn(async (input) => {
      if (input.operation === "prepare") {
        await writeFile(preparedPath, "prepared", "utf8");
        return resultFor(input, { ready: true });
      }
      active++;
      maxActive = Math.max(maxActive, active);
      await new Promise((resolve) => setTimeout(resolve, 20));
      active--;
      if (input.operation === "word" && ++wordAttempts === 1) {
        return { ...resultFor(input), status: "failed" as const, error: "temporary", summary: "temporary" };
      }
      await writeFile(input.outputPath!, input.operation, "utf8");
      return resultFor(input);
    });
    const bridge: OfficeActionBridge = { executeAction };

    const result = await runOfficeWorkflow(bridge, [
      { id: "prepare", app: "excel", action: "edit", operation: "prepare", filePath: sourcePath, outputPath: preparedPath },
      { id: "word", app: "word", action: "insert", operation: "word", filePath: "{{steps.prepare.outputPath}}", outputPath: wordPath, parallelGroup: "reports", retry: { maxAttempts: 2, delayMs: 1 } },
      { id: "ppt", app: "presentation", action: "insert", operation: "ppt", filePath: "{{steps.prepare.outputPath}}", outputPath: pptPath, parallelGroup: "reports" },
      { id: "skip", app: "word", action: "inspect", operation: "should-not-run", filePath: wordPath, when: { step: "prepare", dataPath: "data.ready", equals: false } },
    ], { workflowRoot: path.join(root, "workflows"), transactionRoot: path.join(root, "transactions") });

    expect(result.status).toBe("done");
    expect(result.stepRecords.map((step) => step.status)).toEqual(["done", "done", "done", "skipped"]);
    expect(result.stepRecords[1].attempts).toBe(2);
    expect(result.stepRecords[1].resolvedStep?.filePath).toBe(preparedPath);
    expect(maxActive).toBe(2);
    expect(executeAction.mock.calls.some(([input]) => input.operation === "should-not-run")).toBe(false);
  });

  it("recovers a persisted running workflow after its owner is gone", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-recover-"));
    temporaryRoots.push(root);
    const workflowRoot = path.join(root, "workflows");
    const workflowId = "123e4567-e89b-42d3-a456-426614174000";
    const sourcePath = path.join(root, "source.xlsx");
    await writeFile(sourcePath, "source", "utf8");
    await mkdir(workflowRoot, { recursive: true });
    await writeFile(path.join(workflowRoot, `${workflowId}.json`), JSON.stringify({
      id: workflowId,
      status: "running",
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      leaseExpiresAt: new Date(Date.now() + 60_000).toISOString(),
      steps: [
        { app: "excel", action: "inspect", operation: "first", filePath: sourcePath },
        { app: "word", action: "inspect", operation: "second", filePath: sourcePath },
      ],
      stepRecords: [
        { step: 1, status: "done", artifacts: [], result: resultFor({ app: "excel", action: "inspect", operation: "first", filePath: sourcePath }) },
        { step: 2, status: "running", artifacts: [] },
      ],
      completedSteps: 1,
      nextStep: 2,
    }), "utf8");
    const executeAction = vi.fn(async (input) => resultFor(input));

    const refused = await runOfficeWorkflow({ executeAction }, [], { workflowRoot, workflowId, resume: true });
    const recovered = await runOfficeWorkflow({ executeAction }, [], { workflowRoot, workflowId, resume: true, recoverRunning: true });

    expect(refused.error).toContain("租约过期");
    expect(recovered).toMatchObject({ status: "done", completedSteps: 2 });
    expect(executeAction).toHaveBeenCalledTimes(1);
    expect(executeAction).toHaveBeenCalledWith(expect.objectContaining({ operation: "second" }));
  });

  it("cancels at a safe step boundary without starting the next Office action", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-cancel-"));
    temporaryRoots.push(root);
    const workflowRoot = path.join(root, "workflows");
    const sourcePath = path.join(root, "source.xlsx");
    await writeFile(sourcePath, "source", "utf8");
    let releaseFirst!: () => void;
    let markStarted!: () => void;
    const firstStarted = new Promise<void>((resolve) => { markStarted = resolve; });
    const firstGate = new Promise<void>((resolve) => { releaseFirst = resolve; });
    const executeAction = vi.fn(async (input) => {
      if (input.operation === "first") {
        markStarted();
        await firstGate;
      }
      return resultFor(input);
    });

    const running = runOfficeWorkflow({ executeAction }, [
      { app: "excel", action: "edit", operation: "first", filePath: sourcePath },
      { app: "word", action: "edit", operation: "second", filePath: sourcePath },
    ], { workflowRoot });
    await firstStarted;
    const [record] = await listOfficeWorkflows(workflowRoot);
    await requestOfficeWorkflowCancellation(workflowRoot, record.id);
    releaseFirst();

    const result = await running;

    expect(result).toMatchObject({ status: "cancelled", completedSteps: 1, nextStep: 2 });
    expect(executeAction.mock.calls.map(([input]) => input.operation)).toEqual(["first"]);
    expect((await listOfficeWorkflows(workflowRoot))[0]).toMatchObject({ status: "cancelled" });
  });

  it("prepares dirty open documents before taking transaction snapshots and executing actions", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-prepare-"));
    temporaryRoots.push(root);
    const sourcePath = path.join(root, "source.xlsx");
    await writeFile(sourcePath, "source", "utf8");
    const order: string[] = [];
    const prepareTransaction = vi.fn(async (filePaths: string[]) => { order.push("prepare"); return filePaths; });
    const executeAction = vi.fn(async (input) => { order.push("execute"); return resultFor(input); });

    const result = await runOfficeWorkflow({ executeAction }, [
      { app: "excel", action: "edit", operation: "format", filePath: sourcePath },
    ], { transactionRoot: path.join(root, "transactions"), prepareTransaction });

    expect(result.status).toBe("done");
    expect(order).toEqual(["prepare", "execute"]);
    expect(prepareTransaction).toHaveBeenCalledWith([sourcePath]);
  });

  it("rejects a second executor while the same persisted workflow is still running", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-workflow-lock-"));
    temporaryRoots.push(root);
    const workflowRoot = path.join(root, "workflows");
    const workflowId = "a1111111-1111-4111-8111-111111111111";
    let markStarted!: () => void;
    let releaseStep!: () => void;
    const started = new Promise<void>((resolve) => { markStarted = resolve; });
    const waiting = new Promise<void>((resolve) => { releaseStep = resolve; });
    const bridge: OfficeActionBridge = {
      executeAction: vi.fn(async (input) => {
        markStarted();
        await waiting;
        return resultFor(input);
      }),
    };
    const steps = [{ app: "excel" as const, action: "edit" as const, operation: "format", filePath: "C:/book.xlsx" }];
    const running = runOfficeWorkflow(bridge, steps, { workflowRoot, workflowId });
    await started;

    const duplicate = await runOfficeWorkflow(bridge, steps, { workflowRoot, workflowId });
    expect(duplicate.status).toBe("failed");
    expect(duplicate.error).toContain("正在运行");

    releaseStep();
    await expect(running).resolves.toMatchObject({ status: "done", workflowId });
    expect(bridge.executeAction).toHaveBeenCalledTimes(1);
  });
});

function resultFor(input: any, data?: unknown) {
  return {
    status: "done" as const,
    engine: "com" as const,
    ...input,
    summary: "done",
    changes: input.outputPath ? [{ kind: "create", target: input.outputPath, detail: "created" }] : [],
    data,
  };
}
