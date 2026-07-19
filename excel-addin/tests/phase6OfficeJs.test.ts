import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installStructureExcel } from "./fakes/officeJsStructureFake";

describe("phase6 Office.js structure", () => {
  let gates: ReturnType<typeof installStructureExcel>;

  beforeEach(() => {
    gates = installStructureExcel();
  });
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("reads and writes visibility visible|hidden|veryHidden", async () => {
    const adapter = new OfficeJsAdapter();
    for (const visibility of ["hidden", "veryHidden", "visible"] as const) {
      const set = await adapter.setSheetVisibility("Sheet1", visibility);
      expect(set.ok).toBe(true);
      if (set.ok) expect(set.data.visibility).toBe(visibility);
      const get = await adapter.getSheetVisibility("Sheet1");
      expect(get.ok).toBe(true);
      if (get.ok) expect(get.data.visibility).toBe(visibility);
    }
  });

  it("protects via protect({}, password) second arg; password never in result/state", async () => {
    const adapter = new OfficeJsAdapter();
    const password = "secret-not-persisted";
    const protectedResult = await adapter.protectSheet("Sheet1", password);
    expect(protectedResult.ok).toBe(true);
    if (protectedResult.ok) {
      expect(protectedResult.data.protected).toBe(true);
      expect(JSON.stringify(protectedResult)).not.toContain(password);
    }
    const call = gates.getLastProtectCall();
    expect(call).not.toBeNull();
    expect(call?.options).toEqual({});
    expect(call?.password).toBe(password);
    expect(JSON.stringify(gates.getSheetState("Sheet1"))).not.toContain(password);

    const again = await adapter.protectSheet("Sheet1");
    expect(again.ok).toBe(false);

    const cleared = await adapter.unprotectSheet("Sheet1", password);
    expect(cleared.ok).toBe(true);
    if (cleared.ok) {
      expect(cleared.data.protected).toBe(false);
      expect(JSON.stringify(cleared)).not.toContain(password);
    }
    const unprotectAgain = await adapter.unprotectSheet("Sheet1");
    expect(unprotectAgain.ok).toBe(false);
  });

  it("creates workbook/worksheet named ranges; update formula; rename via delete+add", async () => {
    const adapter = new OfficeJsAdapter();
    const wb = await adapter.createNamedRange({
      name: "SalesTotal",
      refersTo: "Sheet1!$A$1",
      scope: "workbook",
    });
    expect(wb.ok).toBe(true);
    if (wb.ok) {
      expect(wb.data.scope).toBe("workbook");
      expect(wb.data.refersTo).toBe("=Sheet1!$A$1");
    }

    const ws = await adapter.createNamedRange({
      name: "LocalRange",
      refersTo: "=$B$2",
      scope: "worksheet",
      sheetName: "Sheet1",
      visible: true,
    });
    expect(ws.ok).toBe(true);
    if (ws.ok) expect(ws.data.sheetName).toBe("Sheet1");

    const listedWb = await adapter.listNamedRanges({ scope: "workbook" });
    expect(listedWb.ok).toBe(true);
    if (listedWb.ok) expect(listedWb.data.some((item) => item.name === "SalesTotal")).toBe(true);

    const listedWs = await adapter.listNamedRanges({
      scope: "worksheet",
      sheetName: "Sheet1",
    });
    expect(listedWs.ok).toBe(true);
    if (listedWs.ok) expect(listedWs.data.some((item) => item.name === "LocalRange")).toBe(true);

    const formulaOnly = await adapter.updateNamedRange({
      name: "LocalRange",
      scope: "worksheet",
      sheetName: "Sheet1",
      refersTo: "=$C$3",
    });
    expect(formulaOnly.ok).toBe(true);
    if (formulaOnly.ok) {
      expect(formulaOnly.data.name).toBe("LocalRange");
      expect(formulaOnly.data.refersTo).toBe("=$C$3");
    }

    // Rename uses delete+add because NamedItem.name is readonly
    const renamed = await adapter.updateNamedRange({
      name: "SalesTotal",
      scope: "workbook",
      refersTo: "Sheet1!$C$3",
      newName: "SalesTotalV2",
    });
    expect(renamed.ok).toBe(true);
    if (renamed.ok) {
      expect(renamed.data.name).toBe("SalesTotalV2");
      expect(renamed.data.refersTo).toBe("=Sheet1!$C$3");
    }
    const afterRename = await adapter.listNamedRanges({ scope: "workbook" });
    expect(afterRename.ok).toBe(true);
    if (afterRename.ok) {
      expect(afterRename.data.some((item) => item.name === "SalesTotal")).toBe(false);
      expect(afterRename.data.some((item) => item.name === "SalesTotalV2")).toBe(true);
    }

    // Case-insensitive conflict: old name must remain (add-first never runs delete)
    await adapter.createNamedRange({
      name: "KeepMe",
      refersTo: "Sheet1!$D$1",
      scope: "workbook",
    });
    const conflict = await adapter.updateNamedRange({
      name: "SalesTotalV2",
      scope: "workbook",
      newName: "keepme",
    });
    expect(conflict.ok).toBe(false);
    const stillThere = await adapter.listNamedRanges({ scope: "workbook" });
    expect(stillThere.ok).toBe(true);
    if (stillThere.ok) {
      expect(stillThere.data.some((item) => item.name === "SalesTotalV2")).toBe(true);
      expect(stillThere.data.some((item) => item.name === "KeepMe")).toBe(true);
    }

    // Illegal / cell-like name: names.add fails before delete → old name kept
    const illegal = await adapter.updateNamedRange({
      name: "SalesTotalV2",
      scope: "workbook",
      newName: "A1",
    });
    expect(illegal.ok).toBe(false);
    const afterIllegal = await adapter.listNamedRanges({ scope: "workbook" });
    expect(afterIllegal.ok).toBe(true);
    if (afterIllegal.ok) {
      expect(afterIllegal.data.some((item) => item.name === "SalesTotalV2")).toBe(true);
      expect(afterIllegal.data.some((item) => item.name === "A1")).toBe(false);
    }

    const deleted = await adapter.deleteNamedRange({
      name: "SalesTotalV2",
      scope: "workbook",
    });
    expect(deleted.ok).toBe(true);
  });
});


