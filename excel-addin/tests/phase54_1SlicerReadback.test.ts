import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installSlicerExcel } from "./fakes/officeJsSlicerFake";

describe("phase54.1 slicer strict readback", () => {
  let fake: ReturnType<typeof installSlicerExcel>;
  let adapter: OfficeJsAdapter;

  beforeEach(() => {
    fake = installSlicerExcel();
    adapter = new OfficeJsAdapter();
  });
  afterEach(() => {
    fake.uninstall();
  });

  async function createNamed(name: string) {
    const created = await adapter.createSlicer({
      advancedIntent: "interactive-pivot",
      sourceType: "table",
      sourceName: "SalesTable",
      sourceField: "Dept",
      destinationSheet: "Sheet1",
      name,
      caption: "C",
      top: 1,
      left: 2,
      width: 100,
      height: 150,
    });
    expect(created.ok).toBe(true);
    return created;
  }

  it("scalar top string poison ordinary failed", async () => {
    await createNamed("PoisonTop");
    fake.state.poisonScalar("PoisonTop", "top", "10");
    const listTop = await adapter.listSlicers({});
    expect(listTop.ok).toBe(false);
    if (!listTop.ok) {
      expect(listTop.unsupported).not.toBe(true);
      expect(listTop.reason).toMatch(/top/i);
    }
  });

  it("scalar width NaN poison ordinary failed", async () => {
    await createNamed("PoisonWidth");
    fake.state.poisonScalar("PoisonWidth", "width", Number.NaN);
    const listW = await adapter.listSlicers({});
    expect(listW.ok).toBe(false);
    if (!listW.ok) {
      expect(listW.unsupported).not.toBe(true);
      expect(listW.reason).toMatch(/width/i);
    }
  });

  it("scalar isFilterCleared string poison ordinary failed", async () => {
    await createNamed("PoisonBool");
    fake.state.poisonScalar("PoisonBool", "isFilterCleared", "false");
    const listB = await adapter.listSlicers({});
    expect(listB.ok).toBe(false);
    if (!listB.ok) {
      expect(listB.unsupported).not.toBe(true);
      expect(listB.reason).toMatch(/isFilterCleared/i);
    }
  });

  it("item isSelected type poison ordinary failed", async () => {
    await createNamed("PoisonSel");
    fake.state.poisonItem("PoisonSel", 0, "isSelected", "true");
    const g1 = await adapter.getSlicerFilter({ name: "PoisonSel" });
    expect(g1.ok).toBe(false);
    if (!g1.ok) {
      expect(g1.unsupported).not.toBe(true);
      expect(g1.reason).toMatch(/isSelected/i);
    }
  });

  it("item hasData type poison ordinary failed", async () => {
    await createNamed("PoisonHasData");
    fake.state.poisonItem("PoisonHasData", 1, "hasData", 1);
    const g2 = await adapter.getSlicerFilter({ name: "PoisonHasData" });
    expect(g2.ok).toBe(false);
    if (!g2.ok) expect(g2.reason).toMatch(/hasData/i);
  });

  it("item empty key poison ordinary failed", async () => {
    await createNamed("PoisonKey");
    fake.state.poisonItem("PoisonKey", 2, "key", "");
    const g3 = await adapter.getSlicerFilter({ name: "PoisonKey" });
    expect(g3.ok).toBe(false);
    if (!g3.ok) expect(g3.reason).toMatch(/key/i);
  });

  it("selectItems([]) no-op / partial selection ordinary failed", async () => {
    await createNamed("NoOpAll");
    const partial = await adapter.applySlicerFilter({ name: "NoOpAll", keys: ["A"] });
    expect(partial.ok).toBe(true);

    fake.state.setSelectItemsNoOp("NoOpAll");
    const all = await adapter.applySlicerFilter({ name: "NoOpAll", keys: [] });
    expect(all.ok).toBe(false);
    if (!all.ok) {
      expect(all.unsupported).not.toBe(true);
      expect(all.reason).toMatch(/selected|all|isFilterCleared|mismatch/i);
    }
  });

  it("missing getSelectedItems ordinary failed (not verified:false success)", async () => {
    await createNamed("NoGetSel");
    fake.state.removeGetSelectedItems("NoGetSel");
    const got = await adapter.getSlicerFilter({ name: "NoGetSel" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/getSelectedItems/i);
    }
  });

  it("ClientResult.value requires sync before read", async () => {
    await createNamed("CR");
    fake.state.resetClientResultReadFlag();
    const got = await adapter.getSlicerFilter({ name: "CR" });
    expect(got.ok).toBe(true);
    if (got.ok) {
      expect(got.data.selectedKeys.sort()).toEqual(["A", "B", "C"]);
      expect(got.data.verified).toBe(true);
    }
    expect(fake.state.syncCount).toBeGreaterThanOrEqual(2);
    expect(fake.state.lastClientResultReadBeforeSync).toBe(false);
  });

  it("list missing sheetName ordinary failed (not empty success)", async () => {
    await createNamed("OnSheet1");
    const missing = await adapter.listSlicers({ sheetName: "NoSuchSheet" });
    expect(missing.ok).toBe(false);
    if (!missing.ok) {
      expect(missing.unsupported).not.toBe(true);
      expect(missing.reason).toMatch(/sheet|not found/i);
    }
  });

  it("list existing sheet uses worksheet.slicers", async () => {
    await createNamed("OnlySheet1");
    await adapter.createSlicer({
      advancedIntent: "interactive-pivot",
      sourceType: "table",
      sourceName: "SalesTable",
      sourceField: "Dept",
      destinationSheet: "Sheet2",
      name: "OnlySheet2",
    });
    const listed = await adapter.listSlicers({ sheetName: "Sheet1" });
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.slicers.map((s) => s.name)).toEqual(["OnlySheet1"]);
    }
  });

  it("delete fallback without getItemOrNullObject still load/sync items", async () => {
    await createNamed("DelFallback");
    fake.state.disableGetItemOrNullObject();
    const deleted = await adapter.deleteSlicer({ name: "DelFallback" });
    expect(deleted.ok).toBe(true);
    const listed = await adapter.listSlicers({});
    expect(listed.ok).toBe(true);
    if (listed.ok) expect(listed.data.slicers.map((s) => s.name)).not.toContain("DelFallback");
  });

  it("clear verifies all selected + isFilterCleared", async () => {
    await createNamed("ClearMe");
    await adapter.applySlicerFilter({ name: "ClearMe", keys: ["B"] });
    const cleared = await adapter.clearSlicerFilter({ name: "ClearMe" });
    expect(cleared.ok).toBe(true);
    if (cleared.ok) {
      expect(cleared.data.isFilterCleared).toBe(true);
      expect(cleared.data.selectedKeys.sort()).toEqual(["A", "B", "C"]);
      expect(cleared.data.items.every((i) => i.isSelected)).toBe(true);
    }
  });

  it("apply non-empty verifies selectedKeys and item isSelected agreement", async () => {
    await createNamed("ApplySome");
    const applied = await adapter.applySlicerFilter({ name: "ApplySome", keys: ["A", "C"] });
    expect(applied.ok).toBe(true);
    if (applied.ok) {
      expect(applied.data.selectedKeys.sort()).toEqual(["A", "C"]);
      for (const item of applied.data.items) {
        expect(item.isSelected).toBe(["A", "C"].includes(item.key));
      }
    }
  });
});
