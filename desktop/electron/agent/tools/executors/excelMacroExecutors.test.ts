import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type { ExcelVbaBridge, ExcelWorkbookBridge, WpsJsaBridge } from "../contracts/excel";
import { addExcelMacroExecutors } from "./excelMacroExecutors";

const MACRO_TOOLS = ["macro.detect", "macro.run", "macro.write"] as const;

function createMacroExecutors(
  overrides: Partial<{
    workbookBridge: ExcelWorkbookBridge;
    vbaBridge: ExcelVbaBridge;
    jsaBridge: WpsJsaBridge;
  }> = {},
): Map<string, ToolExecutor> {
  const target = new Map<string, ToolExecutor>();
  addExcelMacroExecutors(target, {
    workbookBridge: {
      getHostInfo: vi.fn(),
    } as unknown as ExcelWorkbookBridge,
    vbaBridge: {} as ExcelVbaBridge,
    jsaBridge: {} as WpsJsaBridge,
    ...overrides,
  });
  return target;
}

describe("addExcelMacroExecutors", () => {
  it("registers only macro tools", () => {
    const target = createMacroExecutors();

    expect([...target.keys()]).toEqual([...MACRO_TOOLS]);
    expect(target.has("ui.addControl")).toBe(false);
    expect(target.has("range.write")).toBe(false);
    expect(target.has("workbook.inspect")).toBe(false);
  });

  it("detects host-internal macro languages", async () => {
    const workbookBridge = {
      getHostInfo: vi.fn(async () => ({ host: "wps" as const, version: "12" })),
    } as unknown as ExcelWorkbookBridge;
    const vbaBridge = {
      detectCapabilities: vi.fn(async () => ({
        supported: true,
        host: "wps" as const,
      })),
    } as unknown as ExcelVbaBridge;
    const jsaBridge = {
      detectCapabilities: vi.fn(async () => ({
        language: "javascript" as const,
        supported: true,
        ready: true,
        internal: true as const,
        engine: "WPS JSA" as const,
      })),
    } as unknown as WpsJsaBridge;
    const target = createMacroExecutors({ workbookBridge, vbaBridge, jsaBridge });

    const result = await target.get("macro.detect")!.execute({});

    expect(result).toMatchObject({
      success: true,
      data: {
        host: "wps",
        recommended: "vba",
        available: [
          { language: "vba", ready: true, internal: true, engine: "VBA" },
          {
            language: "javascript",
            ready: true,
            internal: true,
            engine: "WPS JSA",
          },
        ],
      },
    });
  });
});
