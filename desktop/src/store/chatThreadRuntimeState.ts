import type { ThreadMetadata } from "../electronApi";

export interface ReconcileRunningThreadIdsParams {
  threads: ThreadMetadata[];
  runningThreadIds: Record<string, boolean>;
  stoppedThreadIds: Record<string, boolean>;
}

export function isThreadMetadataRunning(
  metadata: Pick<ThreadMetadata, "activeTurnId" | "lastTurnStatus"> | undefined,
): boolean {
  if (!metadata) return false;
  if (metadata.lastTurnStatus) {
    return metadata.lastTurnStatus === "in_progress";
  }
  return Boolean(metadata.activeTurnId);
}

export function reconcileRunningThreadIds({
  threads,
  runningThreadIds,
  stoppedThreadIds,
}: ReconcileRunningThreadIdsParams): Record<string, boolean> {
  return threads.reduce(
    (next, thread) => {
      if (stoppedThreadIds[thread.threadId]) {
        delete next[thread.threadId];
      } else if (isThreadMetadataRunning(thread)) {
        next[thread.threadId] = true;
      } else {
        delete next[thread.threadId];
      }
      return next;
    },
    { ...runningThreadIds },
  );
}
