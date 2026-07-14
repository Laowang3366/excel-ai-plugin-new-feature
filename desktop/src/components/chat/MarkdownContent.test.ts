import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, test } from "vitest";
import {
  MarkdownContent,
  normalizeExternalHttpUrl,
  normalizeVisibleMarkdown,
} from "./MarkdownContent";

describe("MarkdownContent", () => {
  test("renders markdown tables and strong text instead of raw markers", () => {
    const html = renderToStaticMarkup(
      React.createElement(MarkdownContent, {
        content: [
          "**分析报告**",
          "",
          "| 排名 | 姓名 | 总合计 |",
          "|---|---|---:|",
          "| 1 | 王aa | 39,225,778 |",
        ].join("\n"),
      })
    );

    expect(html).toContain("<strong>分析报告</strong>");
    expect(html).toContain("<table>");
    expect(html).not.toContain("**");
    expect(html).not.toContain("|---|");
  });

  test("removes visible heading markers from assistant text", () => {
    const html = renderToStaticMarkup(
      React.createElement(MarkdownContent, {
        content: [
          "### 总逻辑",
          "####下一步",
        ].join("\n"),
      })
    );

    expect(html).toContain("总逻辑");
    expect(html).toContain("下一步");
    expect(html).not.toContain("#");
  });

  test("keeps hash characters inside fenced code blocks", () => {
    const normalized = normalizeVisibleMarkdown([
      "### 总逻辑",
      "```python",
      "# keep this comment",
      "```",
    ].join("\n"));

    expect(normalized).toContain("总逻辑");
    expect(normalized).not.toContain("### 总逻辑");
    expect(normalized).toContain("# keep this comment");
  });

  test("renders external links for system-browser handling only", () => {
    const html = renderToStaticMarkup(
      React.createElement(MarkdownContent, {
        content: "[安全链接](https://example.com/path)",
      })
    );

    expect(html).toContain('href="https://example.com/path"');
    expect(html).toContain('target="_blank"');
    expect(html).toContain('rel="noreferrer noopener"');
    expect(normalizeExternalHttpUrl("javascript:alert(1)")).toBeNull();
    expect(normalizeExternalHttpUrl("file:///C:/secret.txt")).toBeNull();
  });
});
