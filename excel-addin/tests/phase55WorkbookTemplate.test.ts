import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  normalizeRangeAddressForCompare,
  splitSheetQualifiedAddress,
} from "../shared/host/officeJsTemplateReadback";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { WORKBOOK_TEMPLATE_PRESET_STYLES } from "../shared/host/workbookTemplateTypes";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installTemplateExcel } from "./fakes/officeJsTemplateFake";
import { MockHostAdapter } from "./mockHost";

describe("phase55 workbook.template", () => {
  it("registers apply+capture; total 98 unique tools", () => {
    const names = TOOL_DEFINITIONS.map((d) => d.name);
    expect(TOOL_DEFINITIONS).toHaveLength(98);
    expect(new Set(names).size).toBe(98);
    expect(names).toContain("workbook.template.apply");
    expect(names).toContain("workbook.template.capture");
  });

  it("risk, readonly allowlist, sheetNames minItems", () => {
    const apply = TOOL_DEFINITIONS.find((d) => d.name === "workbook.template.apply")!;
    const capture = TOOL_DEFINITIONS.find((d) => d.name === "workbook.template.capture")!;
    expect(apply.riskLevel).toBe("dangerous");
    expect(capture.riskLevel).toBe("safe");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("workbook.template.capture");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("workbook.template.apply");
    const props = (apply.parameters as { properties: { sheetNames: { minItems?: number } } })
      .properties;
    expect(props.sheetNames.minItems).toBe(1);
  });

  it("executor rejects null/unknown/empty sheetNames before Host", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    let hostCalls = 0;
    const orig = host.applyWorkbookTemplate.bind(host);
    host.applyWorkbookTemplate = async (input) => {
      hostCalls += 1;
      return orig(input);
    };

    expect((await ex.execute({ name: "workbook.template.apply", arguments: { extra: true } })).ok).toBe(
      false,
    );
    expect(
      (await ex.execute({ name: "workbook.template.apply", arguments: { preset: null } })).ok,
    ).toBe(false);
    expect(
      (
        await ex.execute({
          name: "workbook.template.apply",
          arguments: { sheetNames: ["Sheet1", "sheet1"] },
        })
      ).ok,
    ).toBe(false);
    const empty = await ex.execute({
      name: "workbook.template.apply",
      arguments: { sheetNames: [] },
    });
    expect(empty.ok).toBe(false);
    if (!empty.ok) expect(empty.error).toMatch(/empty array|minItems|sheetNames/i);
    expect(hostCalls).toBe(0);
  });

  it("MockHost apply defaults and capture", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    const applied = await ex.execute({ name: "workbook.template.apply", arguments: {} });
    expect(applied.ok).toBe(true);
    if (applied.ok) {
      const data = applied.data as {
        preset: string;
        appliedSheets: { readback: { fontName: string; freezeRowCount: number } }[];
      };
      expect(data.preset).toBe("professional");
      expect(data.appliedSheets[0]!.readback.fontName).toBe("微软雅黑");
      expect(data.appliedSheets[0]!.readback.freezeRowCount).toBe(1);
    }
    expect((await ex.execute({ name: "workbook.template.capture", arguments: {} })).ok).toBe(true);
  });

  describe("quote-aware address parse", () => {
    it("splits sheet names containing ! and spaces", () => {
      const a = splitSheetQualifiedAddress("'A!B'!$A$1:$C$2");
      expect(a.sheet).toBe("A!B");
      expect(a.bare.replace(/\$/g, "").toUpperCase()).toBe("A1:C2");
      expect(normalizeRangeAddressForCompare("'Sheet 2'!$B$2:$D$4")).toBe("B2:D4");
      expect(normalizeRangeAddressForCompare("Sheet1!A1")).toBe("A1");
    });
  });

  describe("Office.js apply", () => {
    let gates: ReturnType<typeof installTemplateExcel>;
    beforeEach(() => {
      gates = installTemplateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("applies presets with host readback and plan address match", async () => {
      const adapter = new OfficeJsAdapter();
      for (const preset of ["professional", "financial", "dashboard", "minimal"] as const) {
        gates = installTemplateExcel();
        const result = await adapter.applyWorkbookTemplate({
          preset,
          allSheets: false,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: true,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok).toBe(true);
        if (result.ok) {
          expect(result.data.preset).toBe(preset);
          expect(result.data.appliedSheets[0]!.range).toBe("A1:B2");
          expect(result.data.appliedSheets[0]!.rows).toBe(2);
          expect(result.data.appliedSheets[0]!.columns).toBe(2);
          expect(result.data.appliedSheets[0]!.readback.headerFill).toBe(
            WORKBOOK_TEMPLATE_PRESET_STYLES[preset].headerFill,
          );
          expect(result.data.appliedSheets[0]!.readback.autoFitVerified).toBe(false);
          expect(result.data.limitations.some((l) => l.includes("no add-in-level rollback"))).toBe(
            true,
          );
        }
      }
    });

    it("sheetNames [] / missing / empty intersection fail with zero writes", async () => {
      const adapter = new OfficeJsAdapter();
      const emptyArr = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: [],
        allSheets: true,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(emptyArr.ok).toBe(false);
      expect(gates.writeCalls()).toBe(0);
      expect(gates.excelRunCalls()).toBe(0);

      // force zero run for empty via executor path already covered; host path empty still fails
      gates = installTemplateExcel({ activeSheet: "Sheet1" });
      const noIntersect = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: ["Sheet2"],
        allSheets: false,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(noIntersect.ok).toBe(false);
      if (!noIntersect.ok) {
        expect(noIntersect.unsupported).not.toBe(true);
        expect(noIntersect.reason).toMatch(/empty intersection|no target/i);
      }
      expect(gates.writeCalls()).toBe(0);

      gates = installTemplateExcel();
      const missing = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: ["NoSuch"],
        allSheets: true,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(missing.ok).toBe(false);
      expect(gates.writeCalls()).toBe(0);
    });

    it("empty sheets skipped; freezeRows 0; requirement unsupported", async () => {
      const adapter = new OfficeJsAdapter();
      const withEmpty = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: ["Sheet1", "Empty"],
        allSheets: true,
        fontName: "Arial",
        fontSize: 12,
        autoFit: false,
        showGridlines: true,
        freezeRows: 0,
      });
      expect(withEmpty.ok).toBe(true);
      if (withEmpty.ok) {
        expect(withEmpty.data.appliedSheetCount).toBe(1);
        expect(withEmpty.data.skippedSheets.some((s) => s.name === "Empty")).toBe(true);
        expect(withEmpty.data.appliedSheets[0]!.readback.freezeRowCount).toBe(0);
      }

      for (const opts of [
        { excelApi18: false },
        { missingIsSetSupported: true },
        { isSetSupportedThrows: true },
      ] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const g = installTemplateExcel(opts);
        const result = await adapter.applyWorkbookTemplate({
          preset: "professional",
          allSheets: true,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: true,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
        expect(g.excelRunCalls()).toBe(0);
        expect(g.writeCalls()).toBe(0);
      }
      gates = installTemplateExcel();
    });

    it("surface missing getLocation/freeze → ordinary failed writeCalls===0", async () => {
      const adapter = new OfficeJsAdapter();
      for (const opts of [{ missingFreeze: true }, { missingGetLocation: true }] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const g = installTemplateExcel(opts);
        const result = await adapter.applyWorkbookTemplate({
          preset: "professional",
          allSheets: false,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: true,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
        expect(g.writeCalls()).toBe(0);
      }
      gates = installTemplateExcel();
    });

    it("poison address/rowCount/columnCount/isNullObject/header → ordinary failed", async () => {
      const adapter = new OfficeJsAdapter();
      const cases: Array<Parameters<typeof gates.setPoison>[1]> = [
        { address: 123 as unknown as string },
        { rowCount: 2.5 as unknown as number },
        { rowCount: -1 as unknown as number },
        { columnCount: NaN as unknown as number },
        { columnCount: "2" as unknown as number },
        { headerFill: 123 as unknown as string },
        { isNullObject: "false" as unknown as boolean },
      ];
      for (const poison of cases) {
        gates = installTemplateExcel();
        if (poison && "isNullObject" in poison) {
          gates.setIsNullObject("Sheet1", poison.isNullObject);
        } else {
          gates.setPoison("Sheet1", poison);
        }
        const result = await adapter.applyWorkbookTemplate({
          preset: "professional",
          allSheets: false,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: false,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok, JSON.stringify({ poison, result })).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
      }
    });

    it("post-write address/rowCount/columnCount mismatch fails after writes", async () => {
      const adapter = new OfficeJsAdapter();
      const cases = [
        { address: "Sheet1!Z9:Z10" },
        { rowCount: 99 },
        { columnCount: 77 },
      ] as const;
      for (const poison of cases) {
        gates = installTemplateExcel();
        gates.setPostWritePoison("Sheet1", poison);
        const result = await adapter.applyWorkbookTemplate({
          preset: "professional",
          allSheets: false,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: false,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok, JSON.stringify({ poison, result })).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).not.toBe(true);
          expect(result.reason).toMatch(/mismatch|not a positive integer|address/i);
        }
        expect(gates.writeCalls()).toBeGreaterThan(0);
      }
    });

    it("successful apply ends with sync-before-read false", async () => {
      gates = installTemplateExcel();
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        preset: "professional",
        allSheets: false,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 1,
      });
      expect(result.ok).toBe(true);
      expect(gates.lastClientReadBeforeSync()).toBe(false);
      expect(gates.syncCount()).toBeGreaterThan(0);
    });

    it("quoted sheet name with ! applies address/counts", async () => {
      gates = installTemplateExcel({
        extraSheets: [{ name: "A!B", usedAddress: "A1:C2", rows: 2, cols: 3 }],
        activeSheet: "A!B",
      });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        preset: "financial",
        sheetNames: ["A!B"],
        allSheets: true,
        fontName: "Segoe UI",
        fontSize: 11,
        autoFit: true,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.appliedSheets[0]!.name).toBe("A!B");
        expect(result.data.appliedSheets[0]!.range).toBe("A1:C2");
        expect(result.data.appliedSheets[0]!.rows).toBe(2);
        expect(result.data.appliedSheets[0]!.columns).toBe(3);
        expect(result.data.appliedSheets[0]!.readback.headerRowHeight).toBe(24);
      }
    });
  });

  describe("Office.js capture", () => {
    let gates: ReturnType<typeof installTemplateExcel>;
    beforeEach(() => {
      gates = installTemplateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("single Excel.run multi-sheet shallow snapshot", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.captureWorkbookTemplate();
      expect(result.ok).toBe(true);
      expect(gates.excelRunCalls()).toBe(1);
      if (result.ok) {
        expect(result.data.template.version).toBe(1);
        expect(result.data.sheetCount).toBeGreaterThanOrEqual(3);
        const empty = result.data.template.sheets.find((s) => s.name === "Empty");
        expect(empty?.usedRange).toBeNull();
        expect(empty?.baseStyle).toBeNull();
        const s1 = result.data.template.sheets.find((s) => s.name === "Sheet1");
        expect(s1?.usedRange).toBe("A1:B2");
        expect(s1?.print.orientation).toBe("portrait");
      }
    });

    it("null mixed succeeds with limitations; bad scalars fail", async () => {
      const adapter = new OfficeJsAdapter();
      gates.setBaseNull("Sheet1");
      const mixed = await adapter.captureWorkbookTemplate();
      expect(mixed.ok).toBe(true);
      if (mixed.ok) {
        const s1 = mixed.data.template.sheets.find((s) => s.name === "Sheet1")!;
        expect(s1.baseStyle?.fontName).toBeNull();
        expect(s1.limitations.some((l) => l.includes("mixed/unavailable"))).toBe(true);
      }

      const scalarCases = [
        { baseFontName: 1 },
        { baseFontSize: NaN },
        { baseFontColor: "red" },
        { headerFill: true },
        { headerBold: "yes" },
        { headerRowHeight: Infinity },
        { isNullObject: 0 },
      ] as const;
      for (const poison of scalarCases) {
        gates = installTemplateExcel();
        if ("isNullObject" in poison) {
          gates.setIsNullObject("Sheet1", poison.isNullObject);
        } else {
          gates.setPoison("Sheet1", poison as never);
        }
        const result = await adapter.captureWorkbookTemplate();
        expect(result.ok, JSON.stringify({ poison, result })).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
      }
    });

    it("ExcelApi 1.9 unsupported zero run; >500 resource-limit", async () => {
      const adapter = new OfficeJsAdapter();
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const g = installTemplateExcel({ excelApi19: false });
      const unsup = await adapter.captureWorkbookTemplate();
      expect(unsup.ok).toBe(false);
      if (!unsup.ok) expect(unsup.unsupported).toBe(true);
      expect(g.excelRunCalls()).toBe(0);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installTemplateExcel({ sheetCount: 501 });
      const limited = await adapter.captureWorkbookTemplate();
      expect(limited.ok).toBe(false);
      if (!limited.ok) {
        expect(limited.unsupported).not.toBe(true);
        expect(limited.reason).toMatch(/resource-limit/);
      }
      gates = installTemplateExcel();
    });

    it("quoted sheet with ! captures address", async () => {
      gates = installTemplateExcel({
        extraSheets: [{ name: "A!B", usedAddress: "A1:C2", rows: 2, cols: 3 }],
      });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.captureWorkbookTemplate();
      expect(result.ok).toBe(true);
      expect(gates.excelRunCalls()).toBe(1);
      if (result.ok) {
        const sheet = result.data.template.sheets.find((s) => s.name === "A!B");
        expect(sheet?.usedRange).toBe("A1:C2");
        expect(sheet?.rows).toBe(2);
        expect(sheet?.columns).toBe(3);
      }
    });
  });

  it("WPS typed unsupported for both tools", async () => {
    const wps = new WpsJsaAdapter();
    const apply = await wps.applyWorkbookTemplate({
      preset: "professional",
      allSheets: true,
      fontName: "微软雅黑",
      fontSize: 10.5,
      autoFit: true,
      showGridlines: false,
      freezeRows: 1,
    });
    expect(apply.ok).toBe(false);
    if (!apply.ok) expect(apply.unsupported).toBe(true);
    const capture = await wps.captureWorkbookTemplate();
    expect(capture.ok).toBe(false);
    if (!capture.ok) expect(capture.unsupported).toBe(true);
  });
});
