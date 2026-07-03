import type { RuntimeLongTermMemoryRecord } from "../stateRuntimeTypes";

export function shouldArchiveMemory(
  memory: RuntimeLongTermMemoryRecord,
  now = Date.now(),
): boolean {
  if (memory.status === "archived") return false;
  return memory.expiresAt !== undefined && memory.expiresAt <= now;
}
