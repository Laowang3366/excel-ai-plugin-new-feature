import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools";
import { installShapesExcel } from "./fakes/officeJsShapesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase15 shapes MVP", () => {
  describe("Office.js host", () => {
    let helpers: ReturnType<typeof installShapesExcel>;

    beforeEach(() => {
      helpers = installShapesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("lists default across sheets and multi shapes with writeback", async () => {
      const adapter = new OfficeJsAdapter();
      await adapter.createShape({
        sheetName: "Sheet1",
        kind: "geometric",
        geometricType: "diamond",
        name: "S1A",
      });
      await adapter.createShape({
        sheetName: "Data",
        kind: "textBox",
        text: "D",
        name: "D1",
      });
      const all = await adapter.listShapes();
      expect(all.ok).toBe(true);
      if (all.ok) {
        expect(all.data.map((s) => `${s.sheetName}:${s.name}`).sort()).toEqual([
          "Data:D1",
          "Sheet1:S1A",
        ]);
      }
    });

    it("lists multiple shapes after create with writeback", async () => {
      const adapter = new OfficeJsAdapter();
      const geo = await adapter.createShape({
        sheetName: "Sheet1",
        kind: "geometric",
        geometricType: "rectangle",
        name: "Box1",
        left: 5,
        top: 10,
        width: 120,
        height: 80,
      });
      expect(geo.ok).toBe(true);
      if (geo.ok) {
        expect(geo.data.name).toBe("Box1");
        expect(geo.data.type).toBe("GeometricShape");
        expect(geo.data.geometricShapeType).toBe("Rectangle");
        expect(geo.data.left).toBe(5);
        expect(geo.data.top).toBe(10);
        expect(geo.data.width).toBe(120);
        expect(geo.data.height).toBe(80);
        expect(geo.data.text).toBeNull();
      }

      const box = await adapter.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        text: "Hello",
        name: "TB1",
        left: 1,
        top: 2,
        width: 50,
        height: 20,
      });
      expect(box.ok).toBe(true);
      if (box.ok) {
        expect(box.data.name).toBe("TB1");
        expect(box.data.text).toBe("Hello");
      }

      const listed = await adapter.listShapes("Sheet1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data.map((s) => s.name).sort()).toEqual(["Box1", "TB1"]);
      }
    });

    it("omitted textBox text returns text:null after sync readback", async () => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        name: "EmptyTB",
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        expect(created.data.text).toBeNull();
        expect(created.data.name).toBe("EmptyTB");
      }
    });

    it("does not read text when hasText is false", async () => {
      helpers.seedNoTextShape("Sheet1", "NoText");
      const adapter = new OfficeJsAdapter();
      const listed = await adapter.listShapes("Sheet1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        const row = listed.data.find((s) => s.name === "NoText");
        expect(row?.text).toBeNull();
      }
    });

    it("updates position/size/text/visible with writeback and delete", async () => {
      const adapter = new OfficeJsAdapter();
      await adapter.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        text: "A",
        name: "U1",
      });
      const updated = await adapter.updateShape({
        sheetName: "Sheet1",
        shapeName: "U1",
        left: 30,
        top: 40,
        width: 90,
        height: 70,
        text: "B",
        visible: false,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.left).toBe(30);
        expect(updated.data.top).toBe(40);
        expect(updated.data.width).toBe(90);
        expect(updated.data.height).toBe(70);
        expect(updated.data.text).toBe("B");
        expect(updated.data.visible).toBe(false);
      }
      const renamed = await adapter.updateShape({
        sheetName: "Sheet1",
        shapeName: "U1",
        newName: "U2",
      });
      expect(renamed.ok).toBe(true);
      if (renamed.ok) expect(renamed.data.name).toBe("U2");

      const del = await adapter.deleteShape("Sheet1", "U2");
      expect(del.ok).toBe(true);
      if (del.ok) expect(del.data.deleted).toBe("U2");
      const listed = await adapter.listShapes("Sheet1");
      expect(listed.ok && listed.data.length).toBe(0);
      // Host-observed absence after delete (not input echo alone)
      expect(listed.ok && listed.data.every((s) => s.name !== "U2")).toBe(true);
    });

    it("creates all geometric whitelist types", async () => {
      const adapter = new OfficeJsAdapter();
      for (const geometricType of [
        "rectangle",
        "ellipse",
        "triangle",
        "diamond",
        "rightArrow",
      ] as const) {
        const created = await adapter.createShape({
          sheetName: "Data",
          kind: "geometric",
          geometricType,
          name: `G_${geometricType}`,
        });
        expect(created.ok).toBe(true);
        if (created.ok) {
          expect(String(created.data.geometricShapeType).toLowerCase()).toContain(
            geometricType === "rightArrow" ? "rightarrow" : geometricType,
          );
        }
      }
    });
  });

  describe("executor", () => {
    it("MockHost textBox omit/empty text matches text:null contract", async () => {
      const host = new MockHostAdapter();
      const omitted = await host.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        name: "M1",
      });
      expect(omitted.ok && omitted.data.text).toBeNull();
      const empty = await host.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        name: "M2",
        text: "",
      });
      expect(empty.ok && empty.data.text).toBeNull();
      const filled = await host.createShape({
        sheetName: "Sheet1",
        kind: "textBox",
        name: "M3",
        text: "x",
      });
      expect(filled.ok && filled.data.text).toBe("x");
    });

    it("schema exposes additionalProperties false and positive sizes", async () => {
      const { SHAPE_TOOL_DEFINITIONS } = await import("../shared/tools/shapeDefinitions");
      for (const tool of SHAPE_TOOL_DEFINITIONS) {
        expect(tool.parameters.additionalProperties).toBe(false);
      }
      const create = SHAPE_TOOL_DEFINITIONS.find((t) => t.name === "shape.create")!;
      const props = create.parameters.properties as Record<string, { exclusiveMinimum?: number; minLength?: number }>;
      expect(props.width?.exclusiveMinimum).toBe(0);
      expect(props.height?.exclusiveMinimum).toBe(0);
      expect(props.name?.minLength).toBe(1);
      expect(props.sheetName?.minLength).toBe(1);
    });

    it("accepts legal create/update and rejects invalid", async () => {
      const host = new MockHostAdapter();
      const executor = new ToolExecutor(host);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "geometric",
              geometricType: "ellipse",
              name: "E1",
              width: 10,
              height: 10,
            },
          })
        ).ok,
      ).toBe(true);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: { sheetName: "Sheet1", kind: "geometric" },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "geometric",
              geometricType: "hexagon",
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "textBox",
              width: 0,
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "textBox",
              height: -1,
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "textBox",
              geometricType: null,
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "geometric",
              geometricType: "rectangle",
              name: "   ",
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.create",
            arguments: {
              sheetName: "Sheet1",
              kind: "geometric",
              geometricType: null,
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.update",
            arguments: { sheetName: "Sheet1", shapeName: "E1" },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.update",
            arguments: { sheetName: "Sheet1", shapeName: "E1", newName: "" },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.update",
            arguments: {
              sheetName: "Sheet1",
              shapeName: "E1",
              left: 1,
              unknown: true,
            },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.update",
            arguments: { sheetName: "Sheet1", shapeName: "E1", left: null },
          })
        ).ok,
      ).toBe(false);

      expect(
        (
          await executor.execute({
            name: "shape.update",
            arguments: { sheetName: "Sheet1", shapeName: "E1", left: 9 },
          })
        ).ok,
      ).toBe(true);

      expect(
        (
          await executor.execute({
            name: "shape.list",
            arguments: { sheetName: "Sheet1" },
          })
        ).ok,
      ).toBe(true);

      expect(
        (
          await executor.execute({
            name: "shape.delete",
            arguments: { sheetName: "Sheet1", shapeName: "E1" },
          })
        ).ok,
      ).toBe(true);
    });
  });

  describe("WPS", () => {
    it("returns typed unsupported for all shape tools", async () => {
      const host = new WpsJsaAdapter();
      const executor = new ToolExecutor(host);
      for (const name of [
        "shape.list",
        "shape.create",
        "shape.delete",
        "shape.update",
      ] as const) {
        const result = await executor.execute({
          name,
          arguments:
            name === "shape.list"
              ? {}
              : name === "shape.create"
                ? { sheetName: "S", kind: "textBox" }
                : name === "shape.delete"
                  ? { sheetName: "S", shapeName: "X" }
                  : { sheetName: "S", shapeName: "X", left: 1 },
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
      }
    });
  });
});
