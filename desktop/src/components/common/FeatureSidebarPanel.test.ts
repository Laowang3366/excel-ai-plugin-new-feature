import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, test } from "vitest";
import { FeatureSidebarPanel } from "./FeatureSidebarPanel";

describe("FeatureSidebarPanel", () => {
  test("renders the Chinese feature shortcuts, active option, and form content when open", () => {
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
    expect(html).toContain('aria-selected="true"');
    expect(html).toContain("OCR form");
  });

  test("renders the collapsed empty state without an active option", () => {
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
    expect(html).toContain("选择上方功能开始配置");
    expect(html).not.toContain('aria-selected="true"');
  });
});
