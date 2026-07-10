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
          onClose: () => undefined,
        },
        React.createElement("div", null, "OCR form"),
      ),
    );

    for (const label of ["公式助手", "代码生成", "OCR 识别", "数据清洗", "报告生成", "图表制作"]) {
      expect(html).toContain(label);
    }
    expect(html.match(/<button class="feature-sidebar-shortcut(?: active)?\s*"/g)).toHaveLength(6);
    expect.soft(html).toContain('role="group"');
    expect.soft(html).not.toContain('role="listbox"');
    expect.soft(html).not.toContain('role="option"');
    expect.soft(html).toContain('aria-labelledby="feature-sidebar-title"');
    expect.soft(html).toContain('id="feature-sidebar-title"');
    expect.soft(html.match(/aria-pressed="true"/g) ?? []).toHaveLength(1);
    expect
      .soft(html)
      .toMatch(
        /<button class="feature-sidebar-shortcut active"[^>]*aria-pressed="true"[^>]*>[\s\S]*?<span>OCR 识别<\/span><\/button>/,
      );
    expect(html).toContain("OCR form");
  });

  test("makes the collapsed panel inert without a pressed feature", () => {
    const html = renderToStaticMarkup(
      React.createElement(FeatureSidebarPanel, {
        isOpen: false,
        activeIntent: null,
        language: "zh-CN",
        onIntentClick: () => undefined,
        onClose: () => undefined,
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
        onClose: () => undefined,
      }),
    );

    expect(html).toContain('aria-hidden="false"');
    expect(html).not.toContain('inert=""');
    expect(html).toContain("选择上方功能开始配置");
    expect(html).not.toContain('aria-pressed="true"');
  });
});
