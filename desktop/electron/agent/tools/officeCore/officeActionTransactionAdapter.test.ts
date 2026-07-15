import { describe, expect, it, vi } from "vitest";

import { executeOfficeActionWithTransaction } from "./officeActionTransactionAdapter";
import type { OfficeActionInput, OfficeActionResult } from "./types";

describe("executeOfficeActionWithTransaction", () => {
  it("does not create a nested standalone transaction inside a workflow", async () => {
    const input: OfficeActionInput = {
      app: "excel",
      action: "insert",
      operation: "buildReportPackage",
      filePath: "D:\\reports\\source.xlsx",
      transactionContext: "workflow",
      params: { updateExisting: true },
    };
    const execute = vi.fn(async (): Promise<OfficeActionResult> => doneResult(input));

    const result = await executeOfficeActionWithTransaction(input, {}, execute);

    expect(result.status).toBe("done");
    expect(execute).toHaveBeenCalledWith(input);
  });

  it("rejects a standalone incremental cross-office update before execution without coordination", async () => {
    const input: OfficeActionInput = {
      app: "excel",
      action: "insert",
      operation: "buildReportPackage",
      filePath: "D:\\reports\\source.xlsx",
      params: { updateExisting: true },
    };
    const execute = vi.fn(async (): Promise<OfficeActionResult> => doneResult(input));

    const result = await executeOfficeActionWithTransaction(input, {}, execute);

    expect(result).toMatchObject({
      status: "failed",
      error: "增量跨软件更新需要 Office 事务和文档协调器",
    });
    expect(execute).not.toHaveBeenCalled();
  });
});

function doneResult(input: OfficeActionInput): OfficeActionResult {
  return {
    status: "done",
    engine: "com",
    app: input.app,
    action: input.action,
    operation: input.operation,
    filePath: input.filePath,
    summary: "done",
    changes: [],
  };
}
