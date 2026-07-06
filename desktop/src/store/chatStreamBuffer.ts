import type { AgentEvent } from "../electronApi";
import { ipcApi } from "../services/ipcApi";

export const STREAM_DELTA_STORE_FLUSH_MS = 50;

export type StreamDeltaInput = {
  delta: string;
  itemType: string;
  roundId?: number;
  threadId?: string;
  clientId?: string;
};

export type ChatStreamBufferHandlers = {
  handleAgentEvent: (event: AgentEvent) => void;
  handleStreamDelta: (data: StreamDeltaInput) => void;
};

export function mergeBufferedStreamDeltas(deltas: StreamDeltaInput[]): StreamDeltaInput[] {
  const merged: StreamDeltaInput[] = [];
  for (const delta of deltas) {
    const previous = merged[merged.length - 1];
    if (
      previous &&
      previous.itemType === delta.itemType &&
      previous.roundId === delta.roundId &&
      previous.threadId === delta.threadId &&
      previous.clientId === delta.clientId
    ) {
      previous.delta += delta.delta;
    } else {
      merged.push({ ...delta });
    }
  }
  return merged;
}

let unsubscribeEvent: (() => void) | null = null;
let unsubscribeDelta: (() => void) | null = null;
let pendingStreamDeltas: StreamDeltaInput[] = [];
let streamDeltaFlushTimer: number | null = null;
let activeHandlers: ChatStreamBufferHandlers | null = null;

function flushBufferedStreamDeltas(): void {
  if (streamDeltaFlushTimer) {
    window.clearTimeout(streamDeltaFlushTimer);
    streamDeltaFlushTimer = null;
  }
  if (pendingStreamDeltas.length === 0) return;
  const handlers = activeHandlers;
  if (!handlers) {
    pendingStreamDeltas = [];
    return;
  }
  const batch = mergeBufferedStreamDeltas(pendingStreamDeltas);
  pendingStreamDeltas = [];
  for (const delta of batch) {
    handlers.handleStreamDelta(delta);
  }
}

function scheduleBufferedStreamDeltas(): void {
  if (streamDeltaFlushTimer) return;
  streamDeltaFlushTimer = window.setTimeout(flushBufferedStreamDeltas, STREAM_DELTA_STORE_FLUSH_MS);
}

export function setupChatStreamListeners(handlers: ChatStreamBufferHandlers): void {
  if (typeof window === "undefined") return;

  activeHandlers = handlers;
  flushBufferedStreamDeltas();
  if (unsubscribeEvent) unsubscribeEvent();
  if (unsubscribeDelta) unsubscribeDelta();

  unsubscribeEvent = ipcApi.agent.onEvent((event: AgentEvent) => {
    if (event.type === "stream_delta") {
      pendingStreamDeltas.push(event);
      scheduleBufferedStreamDeltas();
      return;
    }
    flushBufferedStreamDeltas();
    handlers.handleAgentEvent(event);
  });

  unsubscribeDelta = ipcApi.agent.onStreamDelta((data: StreamDeltaInput) => {
    pendingStreamDeltas.push(data);
    scheduleBufferedStreamDeltas();
  });
}
