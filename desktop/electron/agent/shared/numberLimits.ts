export interface ClampNumberOptions {
  fallback: number;
  min: number;
  max: number;
}

export function clampNumber(value: unknown, options: ClampNumberOptions): number {
  const { fallback, min, max } = options;
  if (typeof value !== "number" || !Number.isFinite(value)) return fallback;
  return Math.min(max, Math.max(min, Math.floor(value)));
}
