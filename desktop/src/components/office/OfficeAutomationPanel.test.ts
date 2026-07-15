import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import { OfficeAutomationPanel } from "./OfficeAutomationPanel";
import {
  officeAppLabel,
  parseTemplateVariables,
  shortOfficePath,
} from "./officeAutomationViewModel";

describe("OfficeAutomationPanel", () => {
  it("renders the four direct management views", () => {
    const html = renderToStaticMarkup(React.createElement(OfficeAutomationPanel));
    for (const label of ["文档与对象", "工作流", "事务", "模板"]) expect(html).toContain(label);
    expect(html).toContain('role="tablist"');
    expect(html).toContain('aria-selected="true"');
    expect(html).toContain("没有检测到已打开的 Office 文档");
  });

  it("validates template variables and formats Office labels", () => {
    expect(parseTemplateVariables('{"month":"7月"}')).toEqual({ month: "7月" });
    expect(() => parseTemplateVariables("[]")).toThrow("JSON 对象");
    expect(officeAppLabel("presentation")).toContain("PowerPoint");
    expect(shortOfficePath("C:\\reports\\monthly.docx")).toBe("monthly.docx");
  });
});
