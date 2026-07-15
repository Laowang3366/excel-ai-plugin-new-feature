import type { OfficeActionInput, OfficeActionResult } from "./types";

export type OfficeTransactionStatus = "pending" | "applied" | "undone" | "failed" | "conflicted";

export interface OfficeTransactionSnapshot {
  filePath: string;
  existed: boolean;
  snapshotPath?: string;
  beforeHash?: string;
  afterExisted?: boolean;
  afterSnapshotPath?: string;
  afterHash?: string;
}

export interface OfficeTransactionConflict {
  filePath: string;
  expected: "before" | "after";
  reason: string;
}

export interface OfficeTransactionRestoreFile {
  filePath: string;
  existed: boolean;
  snapshotPath?: string;
}

export interface OfficeTransactionRestoreOptions {
  force?: boolean;
  prepareFiles?: (filePaths: string[]) => Promise<unknown>;
  restoreFiles?: (files: OfficeTransactionRestoreFile[]) => Promise<unknown>;
}

export interface OfficeTransactionRecord {
  id: string;
  workflowId?: string;
  status: OfficeTransactionStatus;
  createdAt: string;
  updatedAt: string;
  steps: OfficeActionInput[];
  results: OfficeActionResult[];
  snapshots: OfficeTransactionSnapshot[];
  artifacts: string[];
  changes: OfficeActionResult["changes"];
  conflicts?: OfficeTransactionConflict[];
  conflictBaseStatus?: Exclude<OfficeTransactionStatus, "conflicted">;
  error?: string;
}
