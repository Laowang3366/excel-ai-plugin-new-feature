import type { OfficeActionBridge } from "../contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "./types";

export async function rollbackOfficeResults(
  bridge: OfficeActionBridge,
  steps: OfficeActionInput[],
  results: OfficeActionResult[],
): Promise<Array<{ step: number; ok: boolean; summary: string }>> {
  const rollback: Array<{ step: number; ok: boolean; summary: string }> = [];
  for (let index = results.length - 1; index >= 0; index--) {
    const transaction = transactionFromResult(results[index]);
    if (!transaction) continue;
    const restored = await bridge.executeAction({
      app: steps[index].app,
      action: "edit",
      operation: "restoreBackup",
      filePath: transaction.sourcePath,
      params: { backupPath: transaction.backupPath },
    });
    rollback.push({ step: index + 1, ok: restored.status === "done", summary: restored.summary });
  }
  return rollback;
}

function transactionFromResult(result: OfficeActionResult): { sourcePath: string; backupPath: string } | undefined {
  if (!result.data || typeof result.data !== "object" || Array.isArray(result.data)) return undefined;
  const transaction = (result.data as Record<string, unknown>).transaction;
  if (!transaction || typeof transaction !== "object" || Array.isArray(transaction)) return undefined;
  const sourcePath = (transaction as Record<string, unknown>).sourcePath;
  const backupPath = (transaction as Record<string, unknown>).backupPath;
  return typeof sourcePath === "string" && typeof backupPath === "string" ? { sourcePath, backupPath } : undefined;
}
