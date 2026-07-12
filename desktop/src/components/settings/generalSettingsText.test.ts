import { describe, expect, it } from "vitest";

import { GENERAL_TEXT } from "./generalSettingsText";

describe("generalSettingsText", () => {
  it("keeps both supported languages available for the general settings view", () => {
    expect(GENERAL_TEXT["zh-CN"].title).toBeTruthy();
    expect(GENERAL_TEXT["en-US"].title).toBe("General");
    expect(GENERAL_TEXT["zh-CN"].dynamicArrayFunctionsEnabled).toBeTruthy();
    expect(GENERAL_TEXT["en-US"].dynamicArrayFunctionsEnabled).toContain("Dynamic array");
  });
});
