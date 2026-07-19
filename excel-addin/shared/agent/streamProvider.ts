export type { AgentStreamProvider, StreamChatRequest } from "./types";

/** Throw if signal is already aborted (DOMException when available). */
export function throwIfAborted(signal?: AbortSignal): void {
  if (!signal?.aborted) return;
  const reason = signal.reason;
  if (reason instanceof Error) throw reason;
  const err =
    typeof DOMException !== "undefined"
      ? new DOMException("The operation was aborted.", "AbortError")
      : Object.assign(new Error("The operation was aborted."), { name: "AbortError" });
  throw err;
}

export function isAbortError(error: unknown): boolean {
  if (!error || typeof error !== "object") return false;
  const name = (error as { name?: string }).name;
  return name === "AbortError";
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
