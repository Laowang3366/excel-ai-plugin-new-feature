import type { ThreadRuntimeSnapshot } from "../../shared/types";

export type IdleThreadUnloadTimer = ReturnType<typeof setTimeout>;

export function clearIdleThreadUnloadTimer(timer: IdleThreadUnloadTimer | null): null {
  if (timer) clearTimeout(timer);
  return null;
}

export function scheduleIdleThreadUnload(input: {
  currentTimer: IdleThreadUnloadTimer | null;
  isRunning: boolean;
  hasActiveThread: boolean;
  getStatus: () => ThreadRuntimeSnapshot;
  sweepIdleThread: () => Promise<boolean>;
  scheduleAgain: () => void;
  now?: () => number;
}): IdleThreadUnloadTimer | null {
  clearIdleThreadUnloadTimer(input.currentTimer);
  if (input.isRunning || !input.hasActiveThread) return null;

  const status = input.getStatus();
  if (status.idleUnloadMs <= 0 || status.lastActiveAt === undefined) return null;

  const now = input.now?.() ?? Date.now();
  const delay = Math.max(0, status.idleUnloadMs - (now - status.lastActiveAt));
  const timer = setTimeout(() => {
    void input.sweepIdleThread().catch(() => {
      input.scheduleAgain();
    });
  }, delay);
  (timer as { unref?: () => void }).unref?.();
  return timer;
}
