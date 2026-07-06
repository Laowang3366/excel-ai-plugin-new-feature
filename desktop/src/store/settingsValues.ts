export const MIN_WINDOW_OPACITY = 0.55;
export const MAX_WINDOW_OPACITY = 1;

export function normalizeWindowOpacity(value: unknown): number {
  const numericValue = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(numericValue)) return MAX_WINDOW_OPACITY;
  const clamped = Math.min(Math.max(numericValue, MIN_WINDOW_OPACITY), MAX_WINDOW_OPACITY);
  return Math.round(clamped * 100) / 100;
}
