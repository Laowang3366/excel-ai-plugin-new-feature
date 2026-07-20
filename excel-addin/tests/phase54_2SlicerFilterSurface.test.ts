import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { installSlicerExcel } from "./fakes/officeJsSlicerFake";

describe("phase54.2 slicer filter surface", () => {
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
    });
    expect(created.ok).toBe(true);
  }

  it("selectedKeys non-string entry ordinary failed", async () => {
    await createNamed("SkNonString");
    fake.state.poisonSelectedKeys("SkNonString", ["A", 1, "C"]);
    const got = await adapter.getSlicerFilter({ name: "SkNonString" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/selectedKeys|string/i);
    }
  });

  it("selectedKeys duplicate ordinary failed", async () => {
    await createNamed("SkDup");
    fake.state.poisonSelectedKeys("SkDup", ["A", "A", "B"]);
    const got = await adapter.getSlicerFilter({ name: "SkDup" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/duplicate/i);
    }
  });

  it("selectedKeys unknown key ordinary failed", async () => {
    await createNamed("SkUnknown");
    fake.state.poisonSelectedKeys("SkUnknown", ["A", "ZZZ"]);
    const got = await adapter.getSlicerFilter({ name: "SkUnknown" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/unknown/i);
    }
  });

  it("duplicate SlicerItem.key ordinary failed", async () => {
    await createNamed("DupItemKey");
    fake.state.poisonItem("DupItemKey", 0, "key", "X");
    fake.state.poisonItem("DupItemKey", 1, "key", "X");
    const got = await adapter.getSlicerFilter({ name: "DupItemKey" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/duplicate key/i);
    }
  });

  it("isFilterCleared true but not all selected ordinary failed", async () => {
    await createNamed("ClearedLie");
    await adapter.applySlicerFilter({ name: "ClearedLie", keys: ["A"] });
    fake.state.poisonScalar("ClearedLie", "isFilterCleared", true);
    const got = await adapter.getSlicerFilter({ name: "ClearedLie" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/isFilterCleared/i);
    }
  });

  it("all selected but isFilterCleared false ordinary failed", async () => {
    await createNamed("AllButFlag");
    // default create is all selected + cleared true; poison flag false
    fake.state.poisonScalar("AllButFlag", "isFilterCleared", false);
    const got = await adapter.getSlicerFilter({ name: "AllButFlag" });
    expect(got.ok).toBe(false);
    if (!got.ok) {
      expect(got.unsupported).not.toBe(true);
      expect(got.reason).toMatch(/isFilterCleared/i);
    }
  });

  it("apply all item keys requires isFilterCleared true", async () => {
    await createNamed("ApplyAllKeys");
    await adapter.applySlicerFilter({ name: "ApplyAllKeys", keys: ["B"] });
    const all = await adapter.applySlicerFilter({
      name: "ApplyAllKeys",
      keys: ["A", "B", "C"],
    });
    expect(all.ok).toBe(true);
    if (all.ok) {
      expect(all.data.isFilterCleared).toBe(true);
      expect(all.data.selectedKeys.sort()).toEqual(["A", "B", "C"]);
    }
  });

  it("apply partial requires isFilterCleared false", async () => {
    await createNamed("ApplyPartial");
    const partial = await adapter.applySlicerFilter({
      name: "ApplyPartial",
      keys: ["A", "C"],
    });
    expect(partial.ok).toBe(true);
    if (partial.ok) {
      expect(partial.data.isFilterCleared).toBe(false);
      expect(partial.data.selectedKeys.sort()).toEqual(["A", "C"]);
    }
  });

  it("ClientResult asserts lastClientResultReadBeforeSync === false", async () => {
    await createNamed("CRFlag");
    fake.state.resetClientResultReadFlag();
    const got = await adapter.getSlicerFilter({ name: "CRFlag" });
    expect(got.ok).toBe(true);
    expect(fake.state.lastClientResultReadBeforeSync).toBe(false);
  });
});
