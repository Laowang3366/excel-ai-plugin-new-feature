import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, test } from "vitest";
import { FeatureSidebarPanel } from "./FeatureSidebarPanel";

describe("FeatureSidebarPanel", () => {
  test("renders six Chinese feature buttons with OCR pressed and active when open", () => {
    const html = renderToStaticMarkup(
      React.createElement(
        FeatureSidebarPanel,
        {
          isOpen: true,
          activeIntent: "ocr",
          language: "zh-CN",
          onIntentClick: () => undefined,
        },
        React.createElement("div", null, "OCR form"),
      ),
    );

    for (const label of [
      "公式助手",
      "代码生成",
      "OCR 识别",
      "数据清洗",
      "报告生成",
      "图表制作",
      "Office 自动化",
    ]) {
      expect(html).toContain(label);
    }
    const shortcutButtons =
      html.match(
        /<button class="feature-sidebar-shortcut feature-[^"]+(?: active)?"[\s\S]*?<\/button>/g,
      ) ?? [];
    expect(shortcutButtons).toHaveLength(7);
    expect.soft(html).toContain('role="group"');
    expect.soft(html).not.toContain('role="listbox"');
    expect.soft(html).not.toContain('role="option"');
    expect.soft(html).toContain('aria-labelledby="feature-sidebar-title"');
    expect.soft(html).toContain('id="feature-sidebar-title"');
    const activeButtons = shortcutButtons.filter((button) =>
      button.includes('aria-pressed="true"'),
    );
    expect(activeButtons).toHaveLength(1);
    expect(activeButtons[0]).toContain('class="feature-sidebar-shortcut feature-ocr active"');
    expect(activeButtons[0]).toContain("<span>OCR 识别</span>");
    expect(html).toContain("OCR form");
    expect(html).not.toContain("feature-sidebar-close");
  });

  test("renders an explicit close command when supplied", () => {
    const html = renderToStaticMarkup(
      React.createElement(FeatureSidebarPanel, {
        isOpen: true,
        activeIntent: "office",
        language: "zh-CN",
        onIntentClick: () => undefined,
        onClose: () => undefined,
      }),
    );
    expect(html).toContain("feature-sidebar-close");
    expect(html).toContain('aria-label="关闭功能模块"');
    expect(html).toContain("office-automation-mode");
  });

  test("makes the collapsed panel inert without a pressed feature", () => {
    const html = renderToStaticMarkup(
      React.createElement(FeatureSidebarPanel, {
        isOpen: false,
        activeIntent: null,
        language: "zh-CN",
        onIntentClick: () => undefined,
      }),
    );

    expect(html).toContain('aria-hidden="true"');
    expect(html).toContain('inert=""');
    expect(html).toContain("选择上方功能开始配置");
    expect(html).not.toContain('aria-pressed="true"');
  });

  test("renders the empty state while the panel is open without an active intent", () => {
    const html = renderToStaticMarkup(
      React.createElement(FeatureSidebarPanel, {
        isOpen: true,
        activeIntent: null,
        language: "zh-CN",
        onIntentClick: () => undefined,
      }),
    );

    expect(html).toContain('aria-hidden="false"');
    expect(html).not.toContain('inert=""');
    expect(html).toContain("选择上方功能开始配置");
    expect(html).not.toContain('aria-pressed="true"');
  });
});
