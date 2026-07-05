import { shouldCompact } from "../../memory/compaction";
import {
  DEFAULT_COMPACTION_CONFIG,
  type CompactionConfig,
  type CompactionReason,
  type Thread,
  type TurnItem,
} from "../../shared/types";
import { buildSessionCompactionConfig } from "./sessionCompactionConfig";

export function buildPreTurnCompactionPlan(input: {
  items: TurnItem[];
  thread: Thread;
  globalConfig?: CompactionConfig;
  pendingReason: CompactionReason | null;
}): { config: CompactionConfig; reason: CompactionReason | null } {
  const globalConfig = input.globalConfig ?? DEFAULT_COMPACTION_CONFIG;
  const sessionContextWindowSize = input.thread.metadata.contextWindowSize
    || globalConfig.contextWindowSize
    || 128_000;
  const config = buildSessionCompactionConfig(globalConfig, sessionContextWindowSize);

  if (input.pendingReason && config.enabled && input.items.length > 0) {
    return { config, reason: input.pendingReason };
  }
  if (shouldCompact(input.items, config)) {
    return { config, reason: "auto_pre_turn" };
  }
  return { config, reason: null };
}
