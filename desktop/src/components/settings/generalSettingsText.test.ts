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
    expect(getWindowOpacityText("zh-CN")).toEqual({
      label: "窗口透明度",
      hint: "降低透明度后，助手窗口会整体半透明，方便查看和操作被遮挡的 Office 内容。",
    });
    expect(getWindowOpacityText("en-US")).toEqual({
      label: "Window opacity",
      hint: "Lower values make the whole assistant window translucent so Office content behind it remains easier to use.",
    });
  });
});
