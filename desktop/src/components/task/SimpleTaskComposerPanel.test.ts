import { describe, expect, it } from "vitest";
import { buildSimpleTaskPayload } from "./SimpleTaskComposerPanel";

describe("buildSimpleTaskPayload", () => {
  it("includes trimmed range and task details when provided", () => {
    expect(
      buildSimpleTaskPayload({
        prefix: "Clean data",
        rangeLabel: "Range",
        requirementLabel: "Requirement",
        range: " A1:C20 ",
        task: " remove blanks ",
      }),
    ).toBe("Clean data\nRange: A1:C20\nRequirement: remove blanks");
  });

  it("omits empty optional details", () => {
    expect(
      buildSimpleTaskPayload({
        prefix: "Create chart",
        rangeLabel: "Range",
        requirementLabel: "Requirement",
        range: " ",
        task: "",
      }),
    ).toBe("Create chart");
  });
});
