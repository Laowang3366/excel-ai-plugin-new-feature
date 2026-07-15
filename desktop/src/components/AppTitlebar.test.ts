import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";

import { getAppText } from "../i18n";
import { AppTitlebar, getOpacityPresentation } from "./AppTitlebar";

describe("AppTitlebar", () => {
  it("renders compact mode, opacity and sidebar state", () => {
    const html = renderToStaticMarkup(
      React.createElement(AppTitlebar, {
        alwaysOnTop: true,
        collapsed: true,
        displayMode: "compact",
        onSetWindowOpacity: vi.fn(),
        onToggleAlwaysOnTop: vi.fn(),
        onToggleCompactMode: vi.fn(),
        onToggleSidebar: vi.fn(),
        showSidebarToggle: true,
        text: getAppText("zh-CN"),
        windowOpacity: 0.7,
      }),
    );

    expect(html).toContain('title="展开侧边栏"');
    expect(html).toContain('aria-label="窗口透明度"');
    expect(html).toContain('value="70"');
    expect(html).toContain("70%");
    expect(html).toContain('title="恢复普通窗口"');
    expect(html).toContain('title="取消置顶"');
  });

  it("calculates the slider presentation from configured opacity bounds", () => {
    expect(getOpacityPresentation(0.55)).toMatchObject({
      fillPercent: 0,
      minPercent: 55,
      maxPercent: 100,
      valuePercent: 55,
    });
    expect(getOpacityPresentation(1)).toMatchObject({
      fillPercent: 100,
      thumbNearValue: false,
      valuePercent: 100,
    });
  });
});
