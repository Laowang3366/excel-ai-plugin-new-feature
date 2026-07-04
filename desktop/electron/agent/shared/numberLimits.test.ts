import { describe, expect, test } from "vitest";
import { clampNumber } from "./numberLimits";

describe("clampNumber", () => {
  test("returns fallback for non-finite values", () => {
    expect(clampNumber(undefined, { fallback: 20, min: 1, max: 100 })).toBe(20);
    expect(clampNumber(Number.NaN, { fallback: 20, min: 1, max: 100 })).toBe(20);
    expect(clampNumber("3", { fallback: 20, min: 1, max: 100 })).toBe(20);
  });

  test("floors finite numbers and clamps them to the configured bounds", () => {
    const options = { fallback: 5, min: 1, max: 10 };

    expect(clampNumber(3.9, options)).toBe(3);
    expect(clampNumber(-1, options)).toBe(1);
    expect(clampNumber(20, options)).toBe(10);
  });
});
