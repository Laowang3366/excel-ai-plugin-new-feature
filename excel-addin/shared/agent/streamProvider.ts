export type { AgentStreamProvider, StreamChatRequest } from "./types";

function makeAbortError(reason?: unknown): Error {
  const message =
    reason instanceof Error
      ? reason.message
      : reason != null && reason !== ""
        ? String(reason)
        : "The operation was aborted.";
  if (typeof DOMException !== "undefined") {
    return new DOMException(message, "AbortError");
  }
  return Object.assign(new Error(message), { name: "AbortError" });
}

/** Normalize any aborted signal (including custom reasons) to AbortError. */
export function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  const reason = signal.reason;
  if (reason instanceof Error && reason.name === "AbortError") throw reason;
  throw makeAbortError(reason);
}

export function isAbortError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  return (error as { name?: string }).name === "AbortError";
}

/** Abortable delay for scripted providers; zero-ms skips timer. */
export function abortableDelay(ms: number, signal?: AbortSignal): Promise<void> {
  throwIfAborted(signal);
  if (ms <= 0) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const onAbort = () => {
      clearTimeout(timer);
      try {
        throwIfAborted(signal);
      } catch (error) {
        reject(error);
      }
    };
    const timer = setTimeout(() => {
      signal?.removeEventListener("abort", onAbort);
      resolve();
    }, ms);
    signal?.addEventListener("abort", onAbort, { once: true });
  });
}
