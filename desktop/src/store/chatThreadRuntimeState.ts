import type { ThreadMetadata } from "../electronApi";

export interface ReconcileRunningThreadIdsParams {
  threads: ThreadMetadata[];
  runningThreadIds: Record<string, boolean>;
  stoppedThreadIds: Record<string, boolean>;
}

export function reconcileRunningThreadIds({
  threads,
  runningThreadIds,
  stoppedThreadIds,
}: ReconcileRunningThreadIdsParams): Record<string, boolean> {
  return threads.reduce((next, thread) => {
    if (stoppedThreadIds[thread.threadId]) {
      delete next[thread.threadId];
    } else if (thread.activeTurnId || thread.lastTurnStatus === "in_progress") {
      next[thread.threadId] = true;
    } else {
      delete next[thread.threadId];
    }
    return next;
  }, { ...runningThreadIds });
}
