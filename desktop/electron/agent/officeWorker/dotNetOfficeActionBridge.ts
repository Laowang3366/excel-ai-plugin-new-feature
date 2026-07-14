import type { OfficeActionBridge } from "../tools/contracts/office";
import type { OfficeActionInput, OfficeActionResult } from "../tools/officeCore/types";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

const DEFAULT_ACTION_TIMEOUT_MS = 120_000;
const DEFAULT_SMOKE_ACTION_TIMEOUT_MS = 30_000;
const MIN_ACTION_TIMEOUT_MS = 5_000;
const MAX_ACTION_TIMEOUT_MS = 600_000;

export class DotNetOfficeActionBridge implements OfficeActionBridge {
  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  executeAction(input: OfficeActionInput): Promise<OfficeActionResult> {
    return this.client.invoke("office.action.execute", { ...input }, actionTimeout(input));
  }
}

export function actionTimeout(input: OfficeActionInput): number {
  const requested = Number(input.params?.actionTimeoutMs);
  if (Number.isFinite(requested)) return clampTimeout(requested);
  if (process.env.WENGGE_OFFICE_SMOKE === "1") {
    const smokeTimeout = Number(process.env.WENGGE_OFFICE_SMOKE_TIMEOUT_MS);
    return clampTimeout(Number.isFinite(smokeTimeout) ? smokeTimeout : DEFAULT_SMOKE_ACTION_TIMEOUT_MS);
  }
  return DEFAULT_ACTION_TIMEOUT_MS;
}

function clampTimeout(timeoutMs: number): number {
  return Math.min(MAX_ACTION_TIMEOUT_MS, Math.max(MIN_ACTION_TIMEOUT_MS, Math.trunc(timeoutMs)));
}
