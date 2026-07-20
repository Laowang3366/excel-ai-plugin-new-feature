import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS } from "../shared/tools";
import { installSlicerExcel } from "./fakes/officeJsSlicerFake";

describe("phase54 slicer Office.js host", () => {
  let fake: ReturnType<typeof installSlicerExcel>;

  beforeEach(() => {
    fake = installSlicerExcel();
  });
  afterEach(() => {
    fake.uninstall();
  });

  it("create/list/update/delete with load+sync readback; no source host fields", async () => {
    const adapter = new OfficeJsAdapter();
    const created = await adapter.createSlicer({
      advancedIntent: "interactive-pivot",
      sourceType: "table",
      sourceName: "SalesTable",
      sourceField: "Dept",
      destinationSheet: "Sheet1",
      name: "DeptSlicer",
      caption: "Department",
      top: 10,
      left: 20,
      width: 120,
      height: 200,
      sortBy: "ascending",
      style: "SlicerStyleLight2",
    });
    expect(created.ok).toBe(true);
    if (!created.ok) return;
    expect(created.data.name).toBe("DeptSlicer");
    expect(created.data.caption).toBe("Department");
    expect(created.data.sheetName).toBe("Sheet1");
    expect(created.data.top).toBe(10);
    expect(created.data.left).toBe(20);
    expect(created.data.width).toBe(120);
    expect(created.data.height).toBe(200);
    expect(created.data.sortBy).toBe("ascending");
    expect(created.data.style).toBe("SlicerStyleLight2");
    expect(created.data.requestedSource).toEqual({
      sourceType: "table",
      sourceName: "SalesTable",
      sourceField: "Dept",
    });
    expect(created.data.limitations?.some((l) => /no source/i.test(l))).toBe(true);
    expect(fake.state.writeCounts.add).toBe(1);
    expect(fake.state.syncCount).toBeGreaterThanOrEqual(3);

    const listed = await adapter.listSlicers({});
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.slicers.map((s) => s.name)).toContain("DeptSlicer");
      expect(listed.data.slicers[0]).not.toHaveProperty("sourceName");
    }

    const updated = await adapter.updateSlicer({
      name: "DeptSlicer",
      caption: "Dept",
      top: 15,
    });
    expect(updated.ok).toBe(true);
    if (updated.ok) {
      expect(updated.data.caption).toBe("Dept");
      expect(updated.data.top).toBe(15);
    }

    const empty = await adapter.updateSlicer({ name: "DeptSlicer" });
    expect(empty.ok).toBe(false);
    if (!empty.ok) expect(empty.reason).toMatch(/empty update/i);

    const deleted = await adapter.deleteSlicer({ name: "DeptSlicer" });
    expect(deleted.ok).toBe(true);
    expect(fake.state.writeCounts.delete).toBe(1);
    const after = await adapter.listSlicers({});
    expect(after.ok).toBe(true);
    if (after.ok) expect(after.data.slicers).toHaveLength(0);
  });

  it("pivot create path; filter apply empty=select all; clear verifies isFilterCleared", async () => {
    const adapter = new OfficeJsAdapter();
    const created = await adapter.createSlicer({
      advancedIntent: "interactive-pivot",
      sourceType: "pivotTable",
      sourceName: "SalesPivot",
      sourceField: "Region",
      destinationSheet: "Sheet2",
      name: "RegionSlicer",
    });
    expect(created.ok).toBe(true);

    const applyAll = await adapter.applySlicerFilter({ name: "RegionSlicer", keys: [] });
    expect(applyAll.ok).toBe(true);
    if (applyAll.ok) {
      expect(applyAll.data.selectedKeys.sort()).toEqual(["A", "B", "C"]);
      expect(applyAll.data.isFilterCleared).toBe(true);
    }

    const applySome = await adapter.applySlicerFilter({
      name: "RegionSlicer",
      keys: ["B"],
    });
    expect(applySome.ok).toBe(true);
    if (applySome.ok) {
      expect(applySome.data.selectedKeys).toEqual(["B"]);
      expect(applySome.data.verified).toBe(true);
      expect(applySome.data.isFilterCleared).toBe(false);
    }

    const got = await adapter.getSlicerFilter({ name: "RegionSlicer" });
    expect(got.ok).toBe(true);
    if (got.ok) expect(got.data.selectedKeys).toEqual(["B"]);

    const cleared = await adapter.clearSlicerFilter({ name: "RegionSlicer" });
    expect(cleared.ok).toBe(true);
    if (cleared.ok) {
      expect(cleared.data.isFilterCleared).toBe(true);
      expect(fake.state.writeCounts.clearFilters).toBe(1);
    }
  });

  it("ExcelApi 1.10 false/missing/throw → typed unsupported, zero Excel.run", async () => {
    const adapter = new OfficeJsAdapter();
    for (const mode of [false, "missing", "throw"] as const) {
      fake.state.setRequirement(mode);
      const before = fake.state.excelRunCalls;
      const result = await adapter.listSlicers({});
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.10|not supported/i);
      }
      expect(fake.state.excelRunCalls).toBe(before);
    }
  });

  it("bad sortBy readback ordinary failed; PropertyNotLoaded without load", async () => {
    const adapter = new OfficeJsAdapter();
    const created = await adapter.createSlicer({
      advancedIntent: "interactive-pivot",
      sourceType: "table",
      sourceName: "SalesTable",
      sourceField: "Dept",
      destinationSheet: "Sheet1",
      name: "BadSort",
      sortBy: "ascending",
    });
    expect(created.ok).toBe(true);
    fake.state.poisonSortBy("BadSort", "Ascending ");
    const listed = await adapter.listSlicers({});
    expect(listed.ok).toBe(false);
    if (!listed.ok) {
      expect(listed.unsupported).not.toBe(true);
      expect(listed.reason).toMatch(/sortBy/i);
    }
  });

  it("WPS typed unsupported for all slicer tools; tool count 98", async () => {
    expect(TOOL_DEFINITIONS).toHaveLength(98);
    const wps = new WpsJsaAdapter();
    for (const call of [
      () => wps.listSlicers({}),
      () =>
        wps.createSlicer({
          advancedIntent: "interactive-pivot",
          sourceType: "table",
          sourceName: "T",
          sourceField: "F",
          destinationSheet: "Sheet1",
        }),
      () => wps.updateSlicer({ name: "X", caption: "Y" }),
      () => wps.deleteSlicer({ name: "X" }),
      () => wps.getSlicerFilter({ name: "X" }),
      () => wps.applySlicerFilter({ name: "X", keys: [] }),
      () => wps.clearSlicerFilter({ name: "X" }),
    ]) {
      const result = await call();
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.host).toBe("wps-jsa");
      }
    }
  });
});
