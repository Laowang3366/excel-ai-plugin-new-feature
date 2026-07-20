/** Align with shared/provider/persistence.ts parseProvider bounds. */
export const CONTEXT_WINDOW_MIN = 1_024;
export const CONTEXT_WINDOW_MAX = 10_000_000;

export type ContextWindowParseOk = { ok: true; value: number };
export type ContextWindowParseErr = { ok: false; error: string };

export function parseContextWindowInput(
  raw: unknown,
): ContextWindowParseOk | ContextWindowParseErr {
  const n =
    typeof raw === "number"
      ? raw
      : typeof raw === "string"
        ? Number(raw.trim())
        : Number.NaN;
  if (!Number.isFinite(n) || !Number.isInteger(n)) {
    return {
      ok: false,
      error: `Context Window 必须是整数（${CONTEXT_WINDOW_MIN}–${CONTEXT_WINDOW_MAX}）`,
    };
  }
  if (n < CONTEXT_WINDOW_MIN || n > CONTEXT_WINDOW_MAX) {
    return {
      ok: false,
      error: `Context Window 须在 ${CONTEXT_WINDOW_MIN}–${CONTEXT_WINDOW_MAX} 之间`,
    };
  }
  return { ok: true, value: n };
}
