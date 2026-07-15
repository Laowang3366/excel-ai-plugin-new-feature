import { describe, expect, it } from "vitest";

import { inspectJsonResourceBudget, type JsonResourceBudget } from "./jsonResourceBudget";

const SMALL_BUDGET: JsonResourceBudget = {
  maxDepth: 2,
  maxNodes: 8,
  maxStringChars: 8,
  maxArrayItems: 3,
  maxObjectProperties: 3,
  maxSerializedBytes: 64,
};

describe("inspectJsonResourceBudget", () => {
  it("accepts bounded JSON values", () => {
    expect(inspectJsonResourceBudget({ name: "sheet", cells: [1, true, null] }, SMALL_BUDGET))
      .toBeNull();
  });

  it("rejects excessive depth, collection sizes and strings", () => {
    expect(inspectJsonResourceBudget({ a: { b: { c: 1 } } }, SMALL_BUDGET)?.message)
      .toContain("嵌套深度");
    expect(inspectJsonResourceBudget([1, 2, 3, 4], SMALL_BUDGET)?.message)
      .toContain("数组");
    expect(inspectJsonResourceBudget({ a: 1, b: 2, c: 3, d: 4 }, SMALL_BUDGET)?.message)
      .toContain("对象字段");
    expect(inspectJsonResourceBudget("123456789", SMALL_BUDGET)?.message)
      .toContain("字符串");
  });

  it("rejects non-JSON and oversized serialized values", () => {
    expect(inspectJsonResourceBudget({ value: Number.NaN }, SMALL_BUDGET)?.message)
      .toContain("有限值");
    expect(inspectJsonResourceBudget({ value: undefined }, SMALL_BUDGET)?.message)
      .toContain("可序列化");
    expect(inspectJsonResourceBudget({ first: "12345678", second: "12345678" }, {
      ...SMALL_BUDGET,
      maxSerializedBytes: 16,
    })?.message).toContain("序列化大小");
    expect(inspectJsonResourceBudget("\n\n\n", {
      ...SMALL_BUDGET,
      maxSerializedBytes: 10,
    })?.message).toContain("序列化大小");
  });
});
