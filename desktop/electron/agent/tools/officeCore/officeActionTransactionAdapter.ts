import path from "node:path";

import type { OfficeDocumentManagerBridge } from "../contracts/office";
import { doneResult, failedResult } from "./results";
import {
  createOfficeBackup,
  listOfficeBackups,
  restoreOfficeBackup,
  type OfficeBackupRecord,
} from "./transactions";
import {
  beginOfficeTransaction,
  finalizeOfficeTransaction,
  listOfficeTransactionPaths,
  recordOfficeTransactionResult,
  undoOfficeTransaction,
} from "./transactionJournal";
import type { OfficeActionInput, OfficeActionResult } from "./types";

export interface OfficeActionTransactionDeps {
  officeDocumentBridge?: OfficeDocumentManagerBridge;
  backupRoot?: string;
  transactionRoot?: string;
}

type OfficeActionExecutor = (input: OfficeActionInput) => Promise<OfficeActionResult>;

export async function executeOfficeActionWithTransaction(
  input: OfficeActionInput,
  deps: OfficeActionTransactionDeps,
  execute: OfficeActionExecutor,
): Promise<OfficeActionResult> {
  if (input.operation === "listBackups") {
    if (!deps.backupRoot) return failedResult(input, "Office 事务备份目录未配置");
    const records = await listOfficeBackups(deps.backupRoot, input.filePath);
    return doneResult({
      engine: "openxml",
      app: input.app,
      action: input.action,
      operation: input.operation,
      filePath: input.filePath,
      summary: `已列出 ${records.length} 个 Office 事务备份`,
      data: { records },
    });
  }

  if (input.operation === "restoreBackup") {
    if (!deps.backupRoot) return failedResult(input, "Office 事务备份目录未配置");
    const backupPath = stringParam(input.params, "backupPath");
    if (!input.filePath || !backupPath) {
      return failedResult(input, "restoreBackup 需要 filePath 和 params.backupPath");
    }
    try {
      await restoreOfficeBackup({
        backupRoot: deps.backupRoot,
        backupPath,
        destinationPath: input.filePath,
      });
      return doneResult({
        engine: "openxml",
        app: input.app,
        action: input.action,
        operation: input.operation,
        filePath: input.filePath,
        summary: "已恢复 Office 事务备份",
        data: { backupPath },
        changes: [
          {
            kind: "transaction-restore",
            target: input.filePath,
            detail: `已从 ${backupPath} 恢复`,
          },
        ],
      });
    } catch (error) {
      return failedResult(input, error);
    }
  }

  if (requiresStandaloneCrossOfficeTransaction(input)) {
    if (!deps.transactionRoot || !deps.officeDocumentBridge) {
      return failedResult(input, "增量跨软件更新需要 Office 事务和文档协调器");
    }
    return executeStandaloneCrossOfficeTransaction(input, {
      transactionRoot: deps.transactionRoot,
      officeDocumentBridge: deps.officeDocumentBridge,
      execute,
    });
  }

  let backup: OfficeBackupRecord | undefined;
  if (deps.backupRoot && shouldCreateBackup(input)) {
    try {
      backup = await createOfficeBackup({
        backupRoot: deps.backupRoot,
        app: input.app,
        operation: input.operation,
        sourcePath: input.filePath!,
      });
    } catch (error) {
      return failedResult(input, error);
    }
  }

  const result = await execute(input);
  if (!backup) return result;

  const data =
    result.data && typeof result.data === "object" && !Array.isArray(result.data)
      ? (result.data as Record<string, unknown>)
      : {};
  return {
    ...result,
    data: { ...data, transaction: backup },
    changes: [
      {
        kind: "transaction-backup",
        target: backup.backupPath,
        detail: `已备份原文件 ${backup.sourcePath}`,
      },
      ...result.changes,
    ],
  };
}

async function executeStandaloneCrossOfficeTransaction(
  input: OfficeActionInput,
  options: {
    transactionRoot: string;
    officeDocumentBridge: OfficeDocumentManagerBridge;
    execute: OfficeActionExecutor;
  },
): Promise<OfficeActionResult> {
  const paths = listOfficeTransactionPaths([input]);
  try {
    await options.officeDocumentBridge.prepareTransaction(paths);
  } catch (error) {
    return failedResult(input, `准备已打开的 Office 文档失败: ${errorMessage(error)}`);
  }

  let transaction: Awaited<ReturnType<typeof beginOfficeTransaction>> | undefined;
  try {
    transaction = await beginOfficeTransaction({ root: options.transactionRoot, steps: [input] });
    const result = await options.execute(input);
    await recordOfficeTransactionResult(options.transactionRoot, transaction, result);
    if (result.status !== "done") {
      await undoOfficeTransaction(
        options.transactionRoot,
        transaction.id,
        transactionRestoreOptions(options.officeDocumentBridge, true),
      );
      return withGroupTransaction(result, transaction.id, "已自动恢复跨软件更新前的文件");
    }
    const completed = await finalizeOfficeTransaction(options.transactionRoot, transaction);
    return withGroupTransaction(result, completed.id, "已创建可整体撤销的跨软件事务");
  } catch (error) {
    let rollbackError = "";
    if (transaction) {
      try {
        await undoOfficeTransaction(
          options.transactionRoot,
          transaction.id,
          transactionRestoreOptions(options.officeDocumentBridge, true),
        );
      } catch (rollback) {
        rollbackError = `；自动恢复失败: ${errorMessage(rollback)}`;
      }
    }
    return failedResult(input, `${errorMessage(error)}${rollbackError}`);
  }
}

function requiresStandaloneCrossOfficeTransaction(input: OfficeActionInput): boolean {
  return (
    input.transactionContext !== "workflow" &&
    input.params?.updateExisting === true &&
    ["exportRangeToWord", "exportRangeToPresentation", "buildReportPackage"].includes(
      input.operation,
    )
  );
}

function transactionRestoreOptions(bridge: OfficeDocumentManagerBridge, force: boolean) {
  return {
    force,
    prepareFiles: (filePaths: string[]) => bridge.prepareTransaction(filePaths),
    restoreFiles: (files: Parameters<OfficeDocumentManagerBridge["restoreTransactionFiles"]>[0]) =>
      bridge.restoreTransactionFiles(files),
  };
}

function withGroupTransaction(
  result: OfficeActionResult,
  transactionId: string,
  detail: string,
): OfficeActionResult {
  const data =
    result.data && typeof result.data === "object" && !Array.isArray(result.data)
      ? (result.data as Record<string, unknown>)
      : {};
  return {
    ...result,
    data: { ...data, transaction: { id: transactionId, kind: "office-group" } },
    changes: [
      { kind: "office-group-transaction", target: transactionId, detail },
      ...result.changes,
    ],
  };
}

function shouldCreateBackup(input: OfficeActionInput): boolean {
  if (
    !input.filePath ||
    input.action === "inspect" ||
    input.action === "validate" ||
    input.action === "snapshot"
  ) {
    return false;
  }
  if (
    [
      "createWorkbook",
      "createDocument",
      "createPresentation",
      "restoreBackup",
      "exportPdf",
      "exportSheetsToPdf",
      "exportHandouts",
      "exportRangeToWord",
      "exportRangeToPresentation",
      "buildReportPackage",
      "traceFormulaDependencies",
      "inspectFormulaDependencies",
      "inspectFormulaBackups",
      "inspectFormulaProtection",
      "inspectPrintSettings",
      "inspectDocumentFormatting",
      "inspectReferences",
      "inspectRevisions",
      "inspectContentControls",
      "mailMerge",
      "batchMailMerge",
      "compareDocuments",
    ].includes(input.operation)
  ) {
    return false;
  }
  if (input.outputPath && path.resolve(input.outputPath) !== path.resolve(input.filePath)) {
    return false;
  }
  return true;
}

function stringParam(params: Record<string, unknown> | undefined, key: string): string | undefined {
  const value = params?.[key];
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
