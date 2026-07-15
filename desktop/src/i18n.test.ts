import { describe, expect, it } from "vitest";

import { getAppText } from "./i18n";

describe("application text resources", () => {
  it("keeps constrained option keys aligned across languages", () => {
    const zh = getAppText("zh-CN");
    const en = getAppText("en-US");

    expect(Object.keys(en.sidebar.intents)).toEqual(Object.keys(zh.sidebar.intents));
    expect(Object.keys(en.chat.permissionLabels)).toEqual(Object.keys(zh.chat.permissionLabels));
    expect(Object.keys(en.chat.simplePlaceholders)).toEqual(
      Object.keys(zh.chat.simplePlaceholders),
    );
    expect(zh.chat.permissionLabels.confirm_all).toBe("高权限自动（危险操作确认）");
    expect(en.sidebar.intents.office).toBe("Office automation");
  });

  it("keeps time formatters and unknown-language fallback behavior", () => {
    expect(getAppText("zh-CN").time.minuteSecond(2, 3)).toBe("2分3秒");
    expect(getAppText("en-US").time.hourParts(1, 2, 3)).toBe("1h 2m 3s");
    expect(getAppText("unsupported" as "zh-CN").app.loading).toBe("Office AI 助手启动中...");
  });
});
