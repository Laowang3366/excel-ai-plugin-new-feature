import { describe, expect, it } from "vitest";
import {
  CONTEXT_WINDOW_MAX,
  CONTEXT_WINDOW_MIN,
  parseContextWindowInput,
} from "../src/providerFormValidation";

describe("parseContextWindowInput", () => {
  it("accepts integers within persistence bounds", () => {
    expect(parseContextWindowInput(CONTEXT_WINDOW_MIN)).toEqual({
      ok: true,
      value: CONTEXT_WINDOW_MIN,
    });
    expect(parseContextWindowInput(128_000)).toEqual({
      ok: true,
      value: 128_000,
    });
    expect(parseContextWindowInput(CONTEXT_WINDOW_MAX)).toEqual({
      ok: true,
      value: CONTEXT_WINDOW_MAX,
    });
  });

  it("rejects non-integers and out-of-range", () => {
    expect(parseContextWindowInput(1024.5).ok).toBe(false);
    expect(parseContextWindowInput(500).ok).toBe(false);
    expect(parseContextWindowInput(CONTEXT_WINDOW_MAX + 1).ok).toBe(false);
    expect(parseContextWindowInput("abc").ok).toBe(false);
    expect(parseContextWindowInput("").ok).toBe(false);
  });
});
