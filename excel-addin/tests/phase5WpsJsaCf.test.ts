import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { oleColorFromHex } from "../shared/host/wpsJsaFormat";
import {
  XL_CELL_VALUE,
  XL_EQUAL,
  XL_GREATER,
} from "../shared/host/wpsJsaValidationConstants";
import {
  installWpsValidationFake,
  uninstallWpsValidationFake,
} from "./fakes/wpsJsaValidationFake";

describe("phase5 WPS JSA conditional formats", () => {
  afterEach(() => {
    uninstallWpsValidationFake();
  });

  describe("member missing → typed unsupported", () => {
    it("FormatConditions missing", async () => {
      installWpsValidationFake({ missingFormatConditions: true });
      const adapter = new WpsJsaAdapter();
      const list = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(list.ok).toBe(false);
      if (!list.ok) expect(list.unsupported).toBe(true);
    });

    it("FormatConditions.Add missing", async () => {
      installWpsValidationFake({ missingFcAdd: true });
      const adapter = new WpsJsaAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE" },
      });
      expect(added.ok).toBe(false);
      if (!added.ok) expect(added.unsupported).toBe(true);
    });
  });

  describe("conditional formats", () => {
    it("lists with 1-based index ids; unsupported host types stay unsupported", async () => {
      const fake = installWpsValidationFake();
      fake.seedCondition("Sheet1", "A1", {
        Type: XL_CELL_VALUE,
        Operator: XL_GREATER,
        Formula1: "0",
      });
      fake.seedCondition("Sheet1", "A1", {
        Type: 3, // color scale-like
        Formula1: "x",
      });
      const adapter = new WpsJsaAdapter();
      const listed = await adapter.listConditionalFormats("Sheet1", "A1");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      expect(listed.data).toHaveLength(2);
      expect(listed.data[0]?.id).toBe("1");
      expect(listed.data[0]?.kind).toBe("cellValue");
      expect(listed.data[0]?.supported).toBe(true);
      expect(listed.data[1]?.id).toBe("2");
      expect(listed.data[1]?.kind).toBe("unsupported");
      expect(listed.data[1]?.supported).toBe(false);
    });

    it("adds cellValue/custom with colors and readback; fixed 1-based id", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const added = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "C1:C5",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "#FF0000",
          fontColor: "#00FF00",
        },
      });
      expect(added.ok).toBe(true);
      if (added.ok) {
        expect(added.data.id).toBe("1");
        expect(added.data.hostType).toBe("CellValue");
      }

      const custom = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "C1:C5",
        rule: { kind: "custom", formula: "=TRUE", fillColor: "#0000FF" },
      });
      expect(custom.ok).toBe(true);
      if (custom.ok) {
        expect(custom.data.id).toBe("2");
        expect(custom.data.kind).toBe("custom");
      }
    });

    it("fails add when host tampers operator/formula/color", async () => {
      installWpsValidationFake({
        tamperCf: { operator: XL_EQUAL, formula1: "99" },
      });
      const bad = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          fillColor: "#FF0000",
        },
      });
      expect(bad.ok).toBe(false);

      installWpsValidationFake({
        tamperCf: { fillColor: oleColorFromHex("#000000")! },
      });
      const badColor = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "equalTo",
          formula1: "1",
          fillColor: "#FF0000",
        },
      });
      expect(badColor.ok).toBe(false);
    });

    it("deletes by 1-based id and verifies remaining multiset", async () => {
      installWpsValidationFake();
      const adapter = new WpsJsaAdapter();
      const a = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=A1>0" },
      });
      const b = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=B1>0" },
      });
      const c = await adapter.addConditionalFormat({
        sheetName: "Sheet1",
        range: "D1",
        rule: { kind: "custom", formula: "=C1>0" },
      });
      expect(a.ok && b.ok && c.ok).toBe(true);
      const deleted = await adapter.deleteConditionalFormat("Sheet1", "D1", "2");
      expect(deleted.ok).toBe(true);
      const listed = await adapter.listConditionalFormats("Sheet1", "D1");
      expect(listed.ok).toBe(true);
      if (!listed.ok) return;
      expect(listed.data.map((x) => x.id)).toEqual(["1", "2"]);
    });

    it("rejects non-between formula2 on add", async () => {
      installWpsValidationFake();
      const bad = await new WpsJsaAdapter().addConditionalFormat({
        sheetName: "Sheet1",
        range: "A1",
        rule: {
          kind: "cellValue",
          operator: "greaterThan",
          formula1: "1",
          formula2: "9",
          fillColor: "#FF0000",
        },
      });
      expect(bad.ok).toBe(false);
    });
  });
});
