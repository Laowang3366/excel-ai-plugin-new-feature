import { describe, expect, test } from "vitest";
import { resolveThinkingBudget } from "./providerClients";

describe("resolveThinkingBudget", () => {
  test("maps reasoning modes to shared thinking budgets", () => {
    expect(resolveThinkingBudget("low")).toBe(5000);
    expect(resolveThinkingBudget("medium")).toBe(10000);
    expect(resolveThinkingBudget("high")).toBe(20000);
    expect(resolveThinkingBudget("max")).toBe(20000);
  });

  test("keeps provider fallback semantics", () => {
    expect(resolveThinkingBudget(undefined)).toBe(20000);
    expect(resolveThinkingBudget("off")).toBe(10000);
  });
});
