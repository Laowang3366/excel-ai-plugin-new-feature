import { describe, expect, it, vi } from "vitest";
import type { OfficeActionInput } from "../tools/officeCore/types";
import { DotNetOfficeActionBridge } from "./dotNetOfficeActionBridge";
import { DotNetOpenXmlBridge } from "./dotNetOpenXmlBridge";
import type { OfficeWorkerClient } from "./officeWorkerClient";

const input: OfficeActionInput = {
  app: "word",
  action: "inspect",
  operation: "inspectReferences",
  filePath: "C:\\docs\\report.docx",
  params: {},
};

describe("DotNetOfficeActionBridge", () => {
  it("forwards the unified action contract to the Worker", async () => {
    const result = { status: "done", engine: "com", changes: [] };
    const invoke = vi.fn().mockResolvedValue(result);
    const bridge = new DotNetOfficeActionBridge({ invoke } as unknown as OfficeWorkerClient);

    await expect(bridge.executeAction(input)).resolves.toBe(result);
    expect(invoke).toHaveBeenCalledWith("office.action.execute", input, 120_000);
  });

  it("clamps caller supplied action timeouts", async () => {
    const invoke = vi.fn().mockResolvedValue({ status: "done" });
    const bridge = new DotNetOfficeActionBridge({ invoke } as unknown as OfficeWorkerClient);

    await bridge.executeAction({ ...input, params: { actionTimeoutMs: 900_000 } });

    expect(invoke).toHaveBeenCalledWith("office.action.execute", expect.any(Object), 600_000);
  });

  it("uses a short default timeout for smoke tests", async () => {
    const invoke = vi.fn().mockResolvedValue({ status: "done" });
    const bridge = new DotNetOfficeActionBridge({ invoke } as unknown as OfficeWorkerClient);
    process.env.WENGGE_OFFICE_SMOKE = "1";
    try {
      await bridge.executeAction(input);
    } finally {
      delete process.env.WENGGE_OFFICE_SMOKE;
    }

    expect(invoke).toHaveBeenCalledWith("office.action.execute", input, 30_000);
  });
});

describe("DotNetOpenXmlBridge", () => {
  it("forwards advanced file actions to the Open XML Worker route", async () => {
    const result = { status: "done", engine: "openxml", changes: [] };
    const invoke = vi.fn().mockResolvedValue(result);
    const bridge = new DotNetOpenXmlBridge({ invoke } as unknown as OfficeWorkerClient);

    await expect(bridge.executeAction!(input)).resolves.toBe(result);
    expect(invoke).toHaveBeenCalledWith("openxml.action.execute", input, 120_000);
  });
});
