import type { AgentTurnCallbacks, AgentTurnInput, Turn } from "../../shared/types";
import type { InputQueue } from "./inputQueue";
import type { ConnectionRequestId, PendingInterruptQueue } from "./pendingInterruptQueue";
import type { TurnState } from "./turnState";

export function enqueueQueuedTurn(input: {
  autoDrainInputQueue: boolean;
  inputQueue: InputQueue;
  turnInput: AgentTurnInput;
  callbacks: AgentTurnCallbacks;
}): { queued: true; queueSize: number } {
  if (!input.autoDrainInputQueue) {
    throw new Error("Agent 正在中断中，请等待停止完成后再发送新请求");
  }
  return {
    queued: true,
    queueSize: input.inputQueue.enqueue({ input: input.turnInput, callbacks: input.callbacks }),
  };
}

export async function interruptCurrentTurn(input: {
  requestId: ConnectionRequestId;
  pendingInterruptQueue: PendingInterruptQueue;
  inputQueue: InputQueue;
  turnState: TurnState;
  disableAutoDrain: () => void;
}): Promise<void> {
  input.pendingInterruptQueue.push(input.requestId);
  input.disableAutoDrain();
  input.inputQueue.clear();
  input.turnState.abortController?.abort();

  try {
    if (input.turnState.turnCompletionPromise) {
      await input.turnState.turnCompletionPromise;
    }
  } finally {
    input.inputQueue.clear();
    input.pendingInterruptQueue.drain();
  }
}

export async function drainQueuedTurns(input: {
  inputQueue: InputQueue;
  isRunning: () => boolean;
  runTurn: (turnInput: AgentTurnInput, callbacks: AgentTurnCallbacks) => Promise<Turn>;
  onTurnError?: (error: unknown) => void;
}): Promise<void> {
  while (!input.isRunning()) {
    const next = input.inputQueue.dequeue();
    if (!next) return;
    try {
      await input.runTurn(next.input, next.callbacks);
    } catch (error) {
      input.onTurnError?.(error);
    }
  }
}

export function shouldRescheduleQueueDrain(input: {
  autoDrainInputQueue: boolean;
  isRunning: boolean;
  queueSize: number;
}): boolean {
  return input.autoDrainInputQueue && !input.isRunning && input.queueSize > 0;
}

export function scheduleQueuedTurnsDrain(input: {
  autoDrainInputQueue: boolean;
  isDrainingInputQueue: boolean;
  isRunning: boolean;
  setDraining: (isDraining: boolean) => void;
  drain: () => Promise<void>;
}): void {
  if (!input.autoDrainInputQueue || input.isDrainingInputQueue || input.isRunning) return;
  input.setDraining(true);
  queueMicrotask(() => {
    void input.drain();
  });
}

export async function drainQueuedTurnsAndReschedule(input: {
  inputQueue: InputQueue;
  isRunning: () => boolean;
  autoDrainInputQueue: () => boolean;
  runTurn: (turnInput: AgentTurnInput, callbacks: AgentTurnCallbacks) => Promise<Turn>;
  setDraining: (isDraining: boolean) => void;
  scheduleDrain: () => void;
  onTurnError?: (error: unknown) => void;
}): Promise<void> {
  try {
    await drainQueuedTurns({
      inputQueue: input.inputQueue,
      isRunning: input.isRunning,
      runTurn: input.runTurn,
      onTurnError: input.onTurnError,
    });
  } finally {
    input.setDraining(false);
    if (
      shouldRescheduleQueueDrain({
        autoDrainInputQueue: input.autoDrainInputQueue(),
        isRunning: input.isRunning(),
        queueSize: input.inputQueue.size(),
      })
    ) {
      input.scheduleDrain();
    }
  }
}
