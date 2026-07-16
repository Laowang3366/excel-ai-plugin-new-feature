import { describe, expect, it } from "vitest";

import {
  officeAppLabel,
  parseTemplateVariables,
  shortOfficePath,
} from "./officeAutomationViewModel";

describe("officeAutomationViewModel", () => {
  it("validates template variables and formats Office labels", () => {
    expect(parseTemplateVariables('{"month":"7月"}')).toEqual({ month: "7月" });
    expect(() => parseTemplateVariables("[]")).toThrow("JSON 对象");
    expect(officeAppLabel("presentation")).toContain("PowerPoint");
    expect(shortOfficePath("C:\\reports\\monthly.docx")).toBe("monthly.docx");
  });
});
