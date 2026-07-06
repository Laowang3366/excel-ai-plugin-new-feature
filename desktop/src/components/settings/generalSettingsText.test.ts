import { describe, expect, it } from "vitest";

import { GENERAL_TEXT, getWindowOpacityText } from "./generalSettingsText";

describe("generalSettingsText", () => {
  it("keeps both supported languages available for the general settings view", () => {
    expect(GENERAL_TEXT["zh-CN"].title).toBeTruthy();
    expect(GENERAL_TEXT["en-US"].title).toBe("General");
    expect(GENERAL_TEXT["zh-CN"].dynamicArrayFunctionsEnabled).toBeTruthy();
    expect(GENERAL_TEXT["en-US"].dynamicArrayFunctionsEnabled).toContain("Dynamic array");
  });

  it("returns localized window opacity labels and hints", () => {
    expect(getWindowOpacityText("zh-CN").label).toBeTruthy();
    expect(getWindowOpacityText("en-US")).toEqual({
      label: "Window opacity",
      hint: "Lower values make the whole assistant window translucent so Office content behind it remains easier to use.",
    });
  });
});
