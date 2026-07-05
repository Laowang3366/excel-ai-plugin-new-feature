import type {
  AgentTurnCallbacks,
  Thread,
  ThreadId,
  ThreadRuntimeSnapshot,
  Turn,
} from "../../shared/types";
import type { createAIClient } from "../../providers/aiClient";
import type { SessionStore } from "../../memory/sessionStore";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import type { LongTermMemoryStore } from "../../memory/longTerm/memoryStore";
import { extractAndWriteTurnMemories } from "../../memory/longTerm/memoryAutoExtraction";

type AgentAIClient = ReturnType<typeof createAIClient>;

export function bindCallbacksToThread(input: {
  callbacks: AgentTurnCallbacks;
  threadId: ThreadId;
  clientId?: string;
}): AgentTurnCallbacks {
  const { callbacks, threadId, clientId } = input;
  return {
    onEvent: (event) => callbacks.onEvent({ ...event, threadId, clientId: event.clientId ?? clientId }),
    onStreamDelta: (delta, itemType, roundId) => {
      callbacks.onStreamDelta?.(delta, itemType, roundId, threadId, clientId);
    },
  };
}

export function attachRolloutEventSink(input: {
  sessionStore: SessionStore;
  stateRuntimeStore?: StateRuntimeStore;
}): void {
  const maybeStore = input.sessionStore as SessionStore & {
    setRolloutEventSink?: SessionStore["setRolloutEventSink"];
  };
  maybeStore.setRolloutEventSink?.(input.stateRuntimeStore ?? null);
}

export async function persistThreadSnapshot(input: {
  stateRuntimeStore?: StateRuntimeStore;
  thread: Thread;
}): Promise<void> {
  if (!input.stateRuntimeStore) return;
  try {
    await input.stateRuntimeStore.upsertThreadSnapshot(input.thread.metadata);
  } catch (error) {
    console.warn("写入线程状态快照失败", error);
  }
}

export async function persistThreadRuntime(input: {
  stateRuntimeStore?: StateRuntimeStore;
  snapshot: ThreadRuntimeSnapshot;
  threadId: ThreadId;
}): Promise<void> {
  if (!input.stateRuntimeStore) return;
  try {
    await input.stateRuntimeStore.updateThreadRuntime({ ...input.snapshot, threadId: input.threadId });
  } catch (error) {
    console.warn("写入线程运行态失败", error);
  }
}

export function scheduleTurnMemoryExtraction(input: {
  aiClient: AgentAIClient;
  memoryStore?: LongTermMemoryStore;
  thread: Thread;
  turn: Turn;
}): void {
  const { aiClient, memoryStore, thread, turn } = input;
  if (!memoryStore || turn.status !== "completed") return;

  void extractAndWriteTurnMemories({
    aiClient,
    memoryStore,
    thread,
    turn,
  }).catch((error) => {
    console.warn("自动写入长期记忆失败:", error);
  });
}
