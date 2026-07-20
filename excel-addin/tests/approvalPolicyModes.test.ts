import { describe, expect, it } from "vitest";
import {
  DEFAULT_PERMISSION_MODE,
  dispositionForRisk,
  isPermissionMode,
  normalizePermissionMode,
  PERMISSION_MODES,
  type PermissionMode,
} from "../shared/agentChat/approvalPolicy";
import { classifyChatTool } from "../shared/agentChat/chatToolPolicy";
import type { RiskLevel } from "../shared/tools/types";

const RISKS: RiskLevel[] = ["safe", "moderate", "dangerous"];

/** mode × risk → expected disposition (desktop-aligned). */
const MATRIX: Record<
  PermissionMode,
  Record<RiskLevel, "direct" | "approval" | "deny">
> = {
  normal: {
    safe: "approval",
    moderate: "approval",
    dangerous: "approval",
  },
  auto_approve_safe: {
    safe: "direct",
    moderate: "approval",
    dangerous: "approval",
  },
  confirm_all: {
    safe: "direct",
    moderate: "direct",
    dangerous: "direct",
  },
};

describe("permission mode matrix", () => {
  it("default is auto_approve_safe (current safe behavior)", () => {
    expect(DEFAULT_PERMISSION_MODE).toBe("auto_approve_safe");
    expect(PERMISSION_MODES).toEqual([
      "normal",
      "auto_approve_safe",
      "confirm_all",
    ]);
  });

  for (const mode of PERMISSION_MODES) {
    for (const risk of RISKS) {
      it(`${mode} × ${risk} → ${MATRIX[mode][risk]}`, () => {
        expect(dispositionForRisk(risk, mode)).toBe(MATRIX[mode][risk]);
      });
    }
  }

  it("unknown risk always denies in every mode (no silent execute)", () => {
    for (const mode of PERMISSION_MODES) {
      expect(dispositionForRisk(undefined, mode)).toBe("deny");
      expect(dispositionForRisk(null, mode)).toBe("deny");
      expect(dispositionForRisk("critical", mode)).toBe("deny");
      expect(dispositionForRisk("", mode)).toBe("deny");
    }
  });

  it("invalid mode falls back to default mapping", () => {
    expect(normalizePermissionMode("nope")).toBe("auto_approve_safe");
    expect(normalizePermissionMode(undefined)).toBe("auto_approve_safe");
    expect(normalizePermissionMode(null)).toBe("auto_approve_safe");
    expect(isPermissionMode("confirm_all")).toBe(true);
    expect(isPermissionMode("full")).toBe(false);
    // invalid mode treated as default → safe direct, moderate approval
    expect(dispositionForRisk("safe", "bogus")).toBe("direct");
    expect(dispositionForRisk("dangerous", "bogus")).toBe("approval");
  });

  it("classifyChatTool respects mode; unknown always deny", () => {
    expect(classifyChatTool("range.read", "normal").disposition).toBe(
      "approval",
    );
    expect(classifyChatTool("range.read", "auto_approve_safe").disposition).toBe(
      "direct",
    );
    expect(classifyChatTool("range.write", "auto_approve_safe").disposition).toBe(
      "approval",
    );
    expect(classifyChatTool("range.write", "confirm_all").disposition).toBe(
      "direct",
    );
    expect(classifyChatTool("nope", "confirm_all").disposition).toBe("deny");
  });
});
