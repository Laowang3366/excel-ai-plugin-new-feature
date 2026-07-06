import { beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../../security/sandbox", () => ({
  evaluateCommand: vi.fn(),
}));

import { evaluateCommand } from "../../security/sandbox";
import { evaluateToolSandboxPolicy } from "./toolSandboxPolicy";

describe("evaluateToolSandboxPolicy", () => {
  beforeEach(() => {
    vi.mocked(evaluateCommand).mockReset();
  });

  it("does not evaluate non-shell tools", async () => {
    const result = await evaluateToolSandboxPolicy("range.read", "{}");

    expect(evaluateCommand).not.toHaveBeenCalled();
    expect(result).toEqual({
      evaluation: null,
      justification: undefined,
      forcedForbidden: false,
      forcedApproval: false,
    });
  });

  it("forces approval and joins prompt rule justifications", async () => {
    const sandboxEvaluation = {
      decision: "prompt",
      evaluation: {
        hits: [
          { matchedPrefix: ["curl"], rule: { decision: "prompt", justification: "Network download" } },
          { matchedPrefix: ["powershell"], rule: { decision: "prompt" } },
          { matchedPrefix: ["Remove-Item"], rule: { decision: "forbidden", justification: "Dangerous delete" } },
        ],
      },
    } as any;
    vi.mocked(evaluateCommand).mockResolvedValueOnce(sandboxEvaluation);

    const result = await evaluateToolSandboxPolicy(
      "shell.execute",
      "{\"command\":\"curl https://example.com\",\"workdir\":\"D:\\\\work\"}"
    );

    expect(evaluateCommand).toHaveBeenCalledWith("curl https://example.com", "D:\\work");
    expect(result).toMatchObject({
      evaluation: sandboxEvaluation,
      justification: "Network download；powershell",
      forcedForbidden: false,
      forcedApproval: true,
    });
  });

  it("falls back to approval when sandbox evaluation fails", async () => {
    vi.mocked(evaluateCommand).mockRejectedValueOnce(new Error("boom"));

    const result = await evaluateToolSandboxPolicy("shell.execute", "{bad json");

    expect(result).toMatchObject({
      evaluation: null,
      justification: "命令策略评估异常，需要用户确认",
      forcedForbidden: false,
      forcedApproval: true,
    });
  });
});
