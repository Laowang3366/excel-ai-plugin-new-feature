import { describe, expect, it } from "vitest";
import { formatEstimatedUsedTokens, formatTokensAsK } from "./modelContextWindows";

describe("modelContextWindows formatting", () => {
  it("shows non-zero usage below 1k as <1k", () => {
    expect(formatEstimatedUsedTokens(0)).toBe("0k");
    expect(formatEstimatedUsedTokens(1)).toBe("<1k");
    expect(formatEstimatedUsedTokens(999)).toBe("<1k");
    expect(formatEstimatedUsedTokens(1000)).toBe("1k");
  });

  it("formats context window size", () => {
    expect(formatTokensAsK(256_000)).toBe("256k");
  });
});
