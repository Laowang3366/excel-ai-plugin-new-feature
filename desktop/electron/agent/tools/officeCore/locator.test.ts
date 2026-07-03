import { describe, expect, it } from "vitest";
import { parseOfficeLocator } from "./locator";

describe("parseOfficeLocator", () => {
  it("parses app-neutral Office locators", () => {
    expect(parseOfficeLocator("range:Sheet1!A1:D10")).toEqual({
      kind: "range",
      value: "Sheet1!A1:D10",
      sheetName: "Sheet1",
      address: "A1:D10",
    });
    expect(parseOfficeLocator("slide:3")).toEqual({ kind: "slide", value: "3", index: 3 });
    expect(parseOfficeLocator("table:1")).toEqual({ kind: "table", value: "1", index: 1 });
  });
});
