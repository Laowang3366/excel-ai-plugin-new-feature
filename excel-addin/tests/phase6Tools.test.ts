import { describe, expect, it } from "vitest";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";

describe("phase6 tools", () => {
  it("registers visibility, protection, namedRange tools", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("sheet.visibility.set");
    expect(names).toContain("sheet.protection.protect");
    expect(names).toContain("namedRange.create");
  });

  it("runs visibility/protection/namedRange success paths", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    const vis = await executor.execute({
      name: "sheet.visibility.set",
      arguments: { sheetName: "Sheet1", visibility: "hidden" },
    });
    expect(vis.ok).toBe(true);

    const password = "mem-only-pwd";
    const protect = await executor.execute({
      name: "sheet.protection.protect",
      arguments: { sheetName: "Sheet1", password },
    });
    expect(protect.ok).toBe(true);
    expect(JSON.stringify(protect)).not.toContain(password);
    const protectAgain = await executor.execute({
      name: "sheet.protection.protect",
      arguments: { sheetName: "Sheet1" },
    });
    expect(protectAgain.ok).toBe(false);

    const unprotect = await executor.execute({
      name: "sheet.protection.unprotect",
      arguments: { sheetName: "Sheet1", password },
    });
    expect(unprotect.ok).toBe(true);
    expect(JSON.stringify(unprotect)).not.toContain(password);

    const created = await executor.execute({
      name: "namedRange.create",
      arguments: {
        name: "A1Name",
        refersTo: "Sheet1!$A$1",
        scope: "workbook",
      },
    });
    expect(created.ok).toBe(true);
  });

  it("rejects invalid visibility/scope/unknown fields/empty name", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    expect(
      (
        await executor.execute({
          name: "sheet.visibility.set",
          arguments: { sheetName: "Sheet1", visibility: "ghost" },
        })
      ).ok,
    ).toBe(false);
    expect(
      (
        await executor.execute({
          name: "namedRange.create",
          arguments: { name: "x", refersTo: "A1", scope: "global" },
        })
      ).ok,
    ).toBe(false);
    expect(
      (
        await executor.execute({
          name: "namedRange.create",
          arguments: { name: "", refersTo: "A1", scope: "workbook" },
        })
      ).ok,
    ).toBe(false);
    expect(
      (
        await executor.execute({
          name: "sheet.protection.protect",
          arguments: { sheetName: "Sheet1", password: "x", options: {} },
        })
      ).ok,
    ).toBe(false);
    // whitespace-only newName is an error, not silent success
    expect(
      (
        await executor.execute({
          name: "namedRange.update",
          arguments: {
            name: "A1Name",
            scope: "workbook",
            newName: "   ",
          },
        })
      ).ok,
    ).toBe(false);
    // workbook scope rejects sheetName
    expect(
      (
        await executor.execute({
          name: "namedRange.list",
          arguments: { scope: "workbook", sheetName: "Sheet1" },
        })
      ).ok,
    ).toBe(false);
    expect(
      (
        await executor.execute({
          name: "namedRange.create",
          arguments: {
            name: "WbOnly",
            refersTo: "Sheet1!$A$1",
            scope: "workbook",
            sheetName: "Sheet1",
          },
        })
      ).ok,
    ).toBe(false);
    // worksheet scope requires sheetName
    expect(
      (
        await executor.execute({
          name: "namedRange.create",
          arguments: { name: "Local", refersTo: "=$A$1", scope: "worksheet" },
        })
      ).ok,
    ).toBe(false);
    // list without scope but with sheetName is invalid
    expect(
      (
        await executor.execute({
          name: "namedRange.list",
          arguments: { sheetName: "Sheet1" },
        })
      ).ok,
    ).toBe(false);
    // update refersTo whitespace-only is invalid
    expect(
      (
        await executor.execute({
          name: "namedRange.update",
          arguments: { name: "A1Name", scope: "workbook", refersTo: "   " },
        })
      ).ok,
    ).toBe(false);
  });

  it("WPS returns unsupported for structure tools", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const call of [
      {
        name: "sheet.visibility.get" as const,
        arguments: { sheetName: "Sheet1" },
      },
      {
        name: "sheet.protection.get" as const,
        arguments: { sheetName: "Sheet1" },
      },
      { name: "namedRange.list" as const, arguments: { scope: "workbook" } },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
