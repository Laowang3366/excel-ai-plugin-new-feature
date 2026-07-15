import { describe, expect, it, vi } from "vitest";

import type { ToolExecutor } from "../../shared/types";
import type { ExcelUiBridge } from "../contracts/excel";
import { addExcelUiExecutors } from "./excelUiExecutors";

const UI_TOOLS = [
  "ui.addControl",
  "ui.removeControl",
  "ui.listControls",
  "ui.createForm",
  "ui.addMenu",
] as const;

function createUiExecutors(
  overrides: Partial<{
    uiBridge: ExcelUiBridge;
  }> = {},
): Map<string, ToolExecutor> {
  const target = new Map<string, ToolExecutor>();
  addExcelUiExecutors(target, {
    uiBridge: {} as ExcelUiBridge,
    ...overrides,
  });
  return target;
}

describe("addExcelUiExecutors", () => {
  it("registers only UI control tools", () => {
    const target = createUiExecutors();

    expect([...target.keys()]).toEqual([...UI_TOOLS]);
    expect(target.has("macro.detect")).toBe(false);
    expect(target.has("range.write")).toBe(false);
    expect(target.has("workbook.open")).toBe(false);
  });

  it("lists sheet controls through the UI bridge", async () => {
    const controls = [{ name: "Button1", controlType: "button" }];
    const uiBridge = {
      listControls: vi.fn(async () => controls),
    } as unknown as ExcelUiBridge;
    const target = createUiExecutors({ uiBridge });

    const result = await target.get("ui.listControls")!.execute({
      sheetName: "Sheet1",
    });

    expect(result).toEqual({ success: true, data: controls });
    expect(uiBridge.listControls).toHaveBeenCalledWith("Sheet1");
  });
});
