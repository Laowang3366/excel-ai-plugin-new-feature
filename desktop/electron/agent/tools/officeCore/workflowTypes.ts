import type { OfficeTransactionRestoreFile } from "./transactionTypes";
import type { OfficeActionInput, OfficeActionResult } from "./types";
import type { OfficeWorkflowStepInput } from "./workflowStepExecution";

export type OfficeWorkflowStatus = "running" | "paused" | "done" | "failed" | "cancelled";

export interface OfficeWorkflowStepRecord {
  step: number;
  id?: string;
  status: "pending" | "running" | "done" | "failed" | "skipped";
  startedAt?: string;
  completedAt?: string;
  attempts?: number;
  artifacts: string[];
  resolvedStep?: OfficeActionInput;
  result?: OfficeActionResult;
}

export interface OfficeWorkflowRecord {
  id: string;
  status: OfficeWorkflowStatus;
  createdAt: string;
  updatedAt: string;
  steps: OfficeWorkflowStepInput[];
  sourceSteps?: OfficeWorkflowStepInput[];
  stepRecords: OfficeWorkflowStepRecord[];
  completedSteps: number;
  nextStep: number;
  transactionId?: string;
  runToken?: string;
  leaseExpiresAt?: string;
  cancelRequested?: boolean;
  variables?: Record<string, unknown>;
  error?: string;
}

export interface OfficeWorkflowResult {
  status: "done" | "failed" | "paused" | "cancelled";
  workflowId?: string;
  transactionId?: string;
  completedSteps: number;
  nextStep?: number;
  failedStep?: number;
  results: OfficeActionResult[];
  stepRecords: OfficeWorkflowStepRecord[];
  rollback: Array<{ step: number; ok: boolean; summary: string }>;
  error?: string;
}

export interface OfficeWorkflowRunOptions {
  workflowRoot?: string;
  transactionRoot?: string;
  workflowId?: string;
  resume?: boolean;
  recoverRunning?: boolean;
  leaseMs?: number;
  failureMode?: "pause" | "rollback";
  cancellationMode?: "pause" | "rollback";
  variables?: Record<string, unknown>;
  prepareTransaction?: (filePaths: string[]) => Promise<unknown>;
  restoreTransaction?: (files: OfficeTransactionRestoreFile[]) => Promise<unknown>;
}
