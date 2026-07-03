import { describe, expect, test } from "vitest";
import { normalize2D } from "./rangeValueUtils";

describe("normalize2D", () => {
  test("keeps existing two-dimensional values", () => {
    const values = [[1, 2], [3, 4]];

    expect(normalize2D(values)).toBe(values);
  });

  test("wraps one-dimensional and scalar values for range writes", () => {
    expect(normalize2D(["A", "B"])).toEqual([["A", "B"]]);
    expect(normalize2D("A1")).toEqual([["A1"]]);
    expect(normalize2D(null)).toEqual([[]]);
  });
});
