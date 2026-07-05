import type { AIClientConfig } from "../../providers/aiClient";
import { createAIClient } from "../../providers/aiClient";
import type {
  CompactionConfig,
  CompactionReason,
  Thread,
} from "../../shared/types";
import {
  createCompactionProvider,
  type CompactionProvider,
} from "./compactionProvider";
import {
  isModelCompHashCompatible,
  resolveModelCompHash,
} from "./modelCompHash";

type AIClient = ReturnType<typeof createAIClient>;

export function applyAIConfigUpdate(input: {
  currentConfig: {
    aiConfig: AIClientConfig;
    compactionConfig?: CompactionConfig;
  };
  nextConfig: AIClientConfig;
  activeThread: Thread | null;
  usesCustomCompactionProvider: boolean;
}): {
  aiClient: AIClient;
  compactionProvider?: CompactionProvider;
  pendingReason: CompactionReason | null;
} {
  const previous = input.currentConfig.aiConfig;
  input.currentConfig.aiConfig = input.nextConfig;
  const aiClient = createAIClient(input.nextConfig);
  const compactionProvider = input.usesCustomCompactionProvider
    ? undefined
    : createCompactionProvider(aiClient, input.currentConfig.compactionConfig);

  if (input.activeThread) {
    input.activeThread.metadata.modelProvider = input.nextConfig.provider;
    input.activeThread.metadata.model = input.nextConfig.model;
    input.activeThread.metadata.contextWindowSize =
      input.nextConfig.contextWindowSize ?? input.activeThread.metadata.contextWindowSize;
    input.activeThread.metadata.compHash = resolveModelCompHash(input.nextConfig);
  }

  let pendingReason: CompactionReason | null = null;
  if (input.activeThread && !isModelCompHashCompatible(previous, input.nextConfig)) {
    pendingReason = "model_changed";
  } else if (input.activeThread && previous.contextWindowSize !== input.nextConfig.contextWindowSize) {
    pendingReason = "context_window_changed";
  }

  return { aiClient, compactionProvider, pendingReason };
}

export function applyCompactionConfigUpdate(input: {
  currentConfig: {
    compactionConfig?: CompactionConfig;
  };
  nextConfig: CompactionConfig;
  aiClient: AIClient;
  activeThread: Thread | null;
  usesCustomCompactionProvider: boolean;
}): {
  compactionProvider?: CompactionProvider;
  pendingReason: CompactionReason | null;
} {
  const previousWindow = input.currentConfig.compactionConfig?.contextWindowSize;
  input.currentConfig.compactionConfig = input.nextConfig;
  const compactionProvider = input.usesCustomCompactionProvider
    ? undefined
    : createCompactionProvider(input.aiClient, input.nextConfig);
  const pendingReason = input.activeThread && previousWindow !== input.nextConfig.contextWindowSize
    ? "context_window_changed"
    : null;
  return { compactionProvider, pendingReason };
}

export function mergePendingCompactionReason(
  current: CompactionReason | null,
  next: CompactionReason | null
): CompactionReason | null {
  if (!next) return current;
  if (current === "model_changed" && next === "context_window_changed") {
    return current;
  }
  return next;
}
