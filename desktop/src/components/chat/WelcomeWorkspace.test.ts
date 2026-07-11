import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, test } from "vitest";
import { WelcomeWorkspace } from "./WelcomeWorkspace";

describe("WelcomeWorkspace", () => {
  test("renders four Chinese quick actions on the empty home page", () => {
    const html = renderToStaticMarkup(
      React.createElement(WelcomeWorkspace, {
        language: "zh-CN",
        onIntentClick: () => undefined,
      }),
    );

    expect(html).toContain("欢迎使用 AI 智能助手");
    for (const label of ["分析数据", "生成代码", "总结文档", "制作图表"]) {
      expect(html).toContain(label);
    }
    expect(html.match(/class="welcome-quick-action /g)).toHaveLength(4);
  });
});
