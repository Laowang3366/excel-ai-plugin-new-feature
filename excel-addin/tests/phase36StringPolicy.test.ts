import { describe, expect, it, vi } from "vitest";
import { ToolExecutor } from "../shared/tools";
import type { HostAdapter } from "../shared/host/types";
import { MockHostAdapter } from "./mockHost";

function run(host: HostAdapter = new MockHostAdapter()) {
  return { host, executor: new ToolExecutor(host) };
}

describe("phase36 string policy (Ident / Value / Clearable)", () => {
  it("Ident: range.read trims sheetName and range before host", async () => {
    const { host, executor } = run();
    const spy = vi.spyOn(host, "readRange");
    expect(
      (
        await executor.execute({
          name: "range.read",
          arguments: { sheetName: " Sheet1 ", range: " A1 " },
        })
      ).ok,
    ).toBe(true);
    expect(spy).toHaveBeenCalledWith("Sheet1", "A1", undefined);
  });

  it("Ident: sheet.rename and chart.delete trim name fields", async () => {
    const { host, executor } = run();
    const rename = vi.spyOn(host, "renameSheet");
    const del = vi.spyOn(host, "deleteChart");
    await executor.execute({
      name: "sheet.rename",
      arguments: { sheetName: " Sheet1 ", newName: " Renamed " },
    });
    expect(rename).toHaveBeenCalledWith("Sheet1", "Renamed");
    await executor.execute({
      name: "chart.delete",
      arguments: { sheetName: " Sheet1 ", chartName: " C1 " },
    });
    expect(del).toHaveBeenCalledWith("Sheet1", "C1");
  });

  it("optionalIdent: empty omit, whitespace fail, padded name trims", async () => {
    const { host, executor } = run();
    const list = vi.spyOn(host, "listTables");
    const create = vi.spyOn(host, "createTable");
    expect((await executor.execute({ name: "table.list", arguments: { sheetName: "" } })).ok).toBe(
      true,
    );
    expect(list).toHaveBeenCalledWith(undefined);
    const blank = await executor.execute({ name: "table.list", arguments: { sheetName: "  " } });
    expect(blank.ok).toBe(false);
    if (!blank.ok) expect(blank.error).toMatch(/sheetName must be non-empty/);
    await executor.execute({
      name: "table.create",
      arguments: { sheetName: "Sheet1", range: "A1:B2", name: " Tbl " },
    });
    expect(create).toHaveBeenCalledWith(
      expect.objectContaining({ sheetName: "Sheet1", address: "A1:B2", name: "Tbl" }),
    );
  });

  it("Ident enums: expand / operation / chartType trim then validate", async () => {
    const { host, executor } = run();
    const read = vi.spyOn(host, "readRange");
    const add = vi.spyOn(host, "addSheet");
    const create = vi.spyOn(host, "createChart");
    await executor.execute({
      name: "range.read",
      arguments: { sheetName: "Sheet1", range: "A1", expand: " none " },
    });
    expect(read).toHaveBeenCalledWith("Sheet1", "A1", "none");
    await executor.execute({
      name: "sheet.operation",
      arguments: { operation: " add ", sheetName: " New " },
    });
    expect(add).toHaveBeenCalledWith("New");
    await executor.execute({
      name: "chart.create",
      arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: " bubble " },
    });
    expect(create).toHaveBeenCalledWith(expect.objectContaining({ chartType: "bubble" }));
  });

  it("Ident: conditionalFormat.delete id trimmed", async () => {
    const { host, executor } = run();
    const spy = vi.spyOn(host, "deleteConditionalFormat");
    await executor.execute({
      name: "conditionalFormat.delete",
      arguments: { sheetName: "Sheet1", range: "A1", id: " cf-1 " },
    });
    expect(spy).toHaveBeenCalledWith("Sheet1", "A1", "cf-1");
  });

  it("structure Ident: visibility/scope/name/sheetName/newName trim", async () => {
    const { host, executor } = run();
    const vis = vi.spyOn(host, "setSheetVisibility");
    const create = vi.spyOn(host, "createNamedRange");
    const update = vi.spyOn(host, "updateNamedRange");
    await executor.execute({
      name: "sheet.visibility.set",
      arguments: { sheetName: " Sheet1 ", visibility: " hidden " },
    });
    expect(vis).toHaveBeenCalledWith("Sheet1", "hidden");
    await executor.execute({
      name: "namedRange.create",
      arguments: {
        name: " N1 ",
        refersTo: "=$A$1",
        scope: " worksheet ",
        sheetName: " Sheet1 ",
      },
    });
    expect(create).toHaveBeenCalledWith(
      expect.objectContaining({ name: "N1", scope: "worksheet", sheetName: "Sheet1" }),
    );
    await executor.execute({
      name: "namedRange.update",
      arguments: { name: " N1 ", scope: " workbook ", newName: " N2 " },
    });
    expect(update).toHaveBeenCalledWith(
      expect.objectContaining({ name: "N1", scope: "workbook", newName: "N2" }),
    );
  });

  it("Value: formula.write keeps padding; whitespace fails", async () => {
    const { host, executor } = run();
    const spy = vi.spyOn(host, "writeFormulas");
    const formula = ' =A1&" x " ';
    expect(
      (
        await executor.execute({
          name: "formula.write",
          arguments: { sheetName: "Sheet1", range: "A1", formula, verify: false },
        })
      ).ok,
    ).toBe(true);
    expect(spy).toHaveBeenCalledWith("Sheet1", "A1", [[formula]]);
    expect(
      (
        await executor.execute({
          name: "formula.write",
          arguments: { sheetName: "Sheet1", range: "A1", formula: "   ", verify: false },
        })
      ).ok,
    ).toBe(false);
  });

  it("Value: namedRange refersTo keeps padding; update blank fails", async () => {
    const { host, executor } = run();
    const create = vi.spyOn(host, "createNamedRange");
    const update = vi.spyOn(host, "updateNamedRange");
    const refersTo = " =$B$2 ";
    await executor.execute({
      name: "namedRange.create",
      arguments: { name: "R1", refersTo, scope: "workbook" },
    });
    expect(create).toHaveBeenCalledWith(expect.objectContaining({ refersTo }));
    await executor.execute({
      name: "namedRange.update",
      arguments: { name: "R1", scope: "workbook", refersTo: " =$C$3 " },
    });
    expect(update).toHaveBeenCalledWith(expect.objectContaining({ refersTo: " =$C$3 " }));
    expect(
      (
        await executor.execute({
          name: "namedRange.update",
          arguments: { name: "R1", scope: "workbook", refersTo: "   " },
        })
      ).ok,
    ).toBe(false);
  });

  it("Value: range.write values and CF/DV rule strings stay raw", async () => {
    const { host, executor } = run();
    const write = vi.spyOn(host, "writeRange");
    const cf = vi.spyOn(host, "addConditionalFormat");
    const dv = vi.spyOn(host, "writeDataValidation");
    await executor.execute({
      name: "range.write",
      arguments: { sheetName: "Sheet1", range: "A1", values: [["  a  "]], verify: false },
    });
    expect(write).toHaveBeenCalledWith("Sheet1", "A1", [["  a  "]]);
    await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: " =A1>0 " },
      },
    });
    expect(cf).toHaveBeenCalledWith(
      expect.objectContaining({ rule: expect.objectContaining({ formula: " =A1>0 " }) }),
    );
    await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "B1",
        rule: { type: "list", listValues: [" a ", "  b"] },
      },
    });
    expect(dv).toHaveBeenCalledWith(
      expect.objectContaining({ rule: expect.objectContaining({ listValues: [" a ", "  b"] }) }),
    );
  });

  it("Clearable: chart.create title raw padding; empty omits", async () => {
    const { host, executor } = run();
    const spy = vi.spyOn(host, "createChart");
    await executor.execute({
      name: "chart.create",
      arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", title: "  Hi  " },
    });
    expect(spy).toHaveBeenCalledWith(expect.objectContaining({ title: "  Hi  " }));
    spy.mockClear();
    await executor.execute({
      name: "chart.create",
      arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", title: "" },
    });
    expect(spy).toHaveBeenCalledWith(expect.objectContaining({ title: undefined }));
  });

  it("Clearable/Special locks: title/headers/shape/tabColor/password", async () => {
    const { host, executor } = run();
    await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
    await host.createShape({
      sheetName: "Sheet1",
      kind: "textBox",
      name: "S1",
      left: 0,
      top: 0,
      width: 100,
      height: 40,
      text: "seed",
    });
    const updateChart = vi.spyOn(host, "updateChart");
    const page = vi.spyOn(host, "setSheetPageLayout");
    const shape = vi.spyOn(host, "updateShape");
    const display = vi.spyOn(host, "setSheetDisplay");
    const protect = vi.spyOn(host, "protectSheet");
    await executor.execute({
      name: "chart.update",
      arguments: { sheetName: "Sheet1", chartName: "C1", title: "  T  " },
    });
    expect(updateChart).toHaveBeenCalledWith(expect.objectContaining({ title: "  T  " }));
    await executor.execute({
      name: "chart.update",
      arguments: { sheetName: "Sheet1", chartName: "C1", title: "" },
    });
    expect(updateChart).toHaveBeenCalledWith(expect.objectContaining({ title: "" }));
    await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: { sheetName: "Sheet1", headers: { left: "  L  ", center: "" } },
    });
    expect(page).toHaveBeenCalledWith(
      expect.objectContaining({ headers: { left: "  L  ", center: "" } }),
    );
    await executor.execute({
      name: "shape.update",
      arguments: { sheetName: "Sheet1", shapeName: "S1", text: "  shape  " },
    });
    expect(shape).toHaveBeenCalledWith(expect.objectContaining({ text: "  shape  " }));
    await executor.execute({
      name: "sheet.display.set",
      arguments: { sheetName: "Sheet1", tabColor: "" },
    });
    expect(display).toHaveBeenCalledWith(expect.objectContaining({ tabColor: "" }));
    await executor.execute({
      name: "sheet.protection.protect",
      arguments: { sheetName: "Sheet1", password: "  secret  " },
    });
    expect(protect).toHaveBeenCalledWith("Sheet1", "  secret  ");
  });

  it("unknown still before host; pageLayout unknown unchanged", async () => {
    const { host, executor } = run();
    const read = vi.spyOn(host, "readRange");
    const unknown = await executor.execute({
      name: "range.read",
      arguments: { sheetName: "Sheet1", range: "A1", __unknown: 1 },
    });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.error).toMatch(/unknown field: __unknown/);
    expect(read).not.toHaveBeenCalled();
    const pl = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: { sheetName: "Sheet1", __unknown: true },
    });
    expect(pl.ok).toBe(false);
    if (!pl.ok) expect(pl.error).toMatch(/unknown field: __unknown/);
  });
});
