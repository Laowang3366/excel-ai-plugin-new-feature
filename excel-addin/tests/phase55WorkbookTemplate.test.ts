import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  normalizeRangeAddressForCompare,
  requireHexColor,
  splitSheetQualifiedAddress,
} from "../shared/host/officeJsTemplateReadback";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { WORKBOOK_TEMPLATE_PRESET_STYLES } from "../shared/host/workbookTemplateTypes";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installTemplateExcel } from "./fakes/officeJsTemplateFake";
import { MockHostAdapter } from "./mockHost";

const APPLY_DEFAULT = {
  preset: "professional" as const,
  allSheets: false,
  fontName: "微软雅黑",
  fontSize: 10.5,
  autoFit: false,
  showGridlines: false,
  freezeRows: 1,
};

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
    const empty = await ex.execute({
      name: "workbook.template.apply",
      arguments: { sheetNames: [] },
    });
    expect(empty.ok).toBe(false);
    expect(hostCalls).toBe(0);
  });

  it("MockHost apply defaults and capture", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    const applied = await ex.execute({ name: "workbook.template.apply", arguments: {} });
    expect(applied.ok).toBe(true);
    expect((await ex.execute({ name: "workbook.template.capture", arguments: {} })).ok).toBe(true);
  });

  it("requireHexColor rejects missing # prefix", () => {
    expect(() => requireHexColor("1F4E79", "c")).toThrow(/#RRGGBB/);
    expect(requireHexColor("#1f4e79", "c")).toBe("#1F4E79");
  });

  describe("quote-aware address parse", () => {
    it("splits sheet names containing ! and spaces", () => {
      const a = splitSheetQualifiedAddress("'A!B'!$A$1:$C$2");
      expect(a.sheet).toBe("A!B");
      expect(normalizeRangeAddressForCompare("'Sheet 2'!$B$2:$D$4")).toBe("B2:D4");
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
        const result = await adapter.applyWorkbookTemplate({ ...APPLY_DEFAULT, preset, autoFit: true });
        expect(result.ok).toBe(true);
        if (result.ok) {
          expect(result.data.appliedSheets[0]!.range).toBe("A1:B2");
          expect(result.data.appliedSheets[0]!.readback.headerFill).toBe(
            WORKBOOK_TEMPLATE_PRESET_STYLES[preset].headerFill,
          );
          expect(result.data.limitations.some((l) => l.includes("no add-in-level rollback"))).toBe(
            true,
          );
        }
      }
      expect(gates.everReadBeforeSync()).toBe(false);
    });

    it("multi-cell apply never bulk-loads text/values", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate(APPLY_DEFAULT);
      expect(result.ok).toBe(true);
      expect(gates.bulkTextReadCalls()).toBe(0);
      expect(gates.bulkValuesReadCalls()).toBe(0);
    });

    it("1x1 blank cell skips without treating multi-cell as empty via text", async () => {
      gates = installTemplateExcel({
        extraSheets: [{ name: "Blank1", usedAddress: "A1", rows: 1, cols: 1, text: "" }],
        activeSheet: "Blank1",
      });
      // ensure not isNull empty flag
      gates.setEmpty("Blank1", false);
      gates.setUsedDims("Blank1", 1, 1, "A1", "");
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        ...APPLY_DEFAULT,
        sheetNames: ["Blank1"],
        allSheets: true,
        freezeRows: 0,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.skippedSheets.some((s) => s.name === "Blank1")).toBe(true);
        expect(result.data.appliedSheetCount).toBe(0);
      }
      // single-cell may load text once
      expect(gates.bulkTextReadCalls()).toBeGreaterThanOrEqual(1);
    });

    it("sheetNames [] / empty intersection fail with zero writes", async () => {
      const adapter = new OfficeJsAdapter();
      const emptyArr = await adapter.applyWorkbookTemplate({
        ...APPLY_DEFAULT,
        sheetNames: [],
        allSheets: true,
      });
      expect(emptyArr.ok).toBe(false);
      expect(gates.writeCalls()).toBe(0);
      expect(gates.excelRunCalls()).toBe(0);

      gates = installTemplateExcel({ activeSheet: "Sheet1" });
      const noIntersect = await adapter.applyWorkbookTemplate({
        ...APPLY_DEFAULT,
        sheetNames: ["Sheet2"],
        allSheets: false,
      });
      expect(noIntersect.ok).toBe(false);
      if (!noIntersect.ok) expect(noIntersect.reason).toMatch(/no target/i);
      expect(gates.writeCalls()).toBe(0);
    });

    it("surface missing members → ordinary failed writeCalls===0", async () => {
      const adapter = new OfficeJsAdapter();
      const cases = [
        { missingFreeze: true },
        { missingGetLocation: true },
        { missingUsedRange: true },
        { missingGetRange: true },
        { missingRangeLoad: true },
        { missingFormat: true },
        { missingFormatLoad: true },
        { missingFont: true },
        { missingFontLoad: true },
        { missingFill: true },
        { missingFillLoad: true },
        { missingAutofit: true },
      ] as const;
      for (const opts of cases) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const g = installTemplateExcel(opts);
        const result = await adapter.applyWorkbookTemplate({
          ...APPLY_DEFAULT,
          autoFit: "missingAutofit" in opts,
        });
        expect(result.ok, JSON.stringify({ opts, result })).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
        expect(g.writeCalls()).toBe(0);
      }
      // Header does not require autofit when autoFit true: only UsedRange
      gates = installTemplateExcel();
      const okAuto = await adapter.applyWorkbookTemplate({ ...APPLY_DEFAULT, autoFit: true });
      expect(okAuto.ok).toBe(true);
    });

    it("table-driven poison for apply readback fields", async () => {
      const adapter = new OfficeJsAdapter();
      const poisons: Array<Record<string, unknown>> = [
        { fontName: 1 },
        { fontSize: "10" },
        { headerFontName: "" },
        { headerFontSize: NaN },
        { headerFill: "1F4E79" }, // missing #
        { headerFontColor: true },
        { headerBold: "yes" },
        { headerAlignment: "Left" },
        { headerWrap: 1 },
        { headerRowHeight: 0 },
        { showGridlines: "false" },
        { isNullObject: "false" },
        { freezeAddress: 123 },
        { freezeColumnCount: -1 },
        { address: 123 },
        { rowCount: 2.5 },
        { columnCount: "2" },
      ];
      for (const poison of poisons) {
        gates = installTemplateExcel();
        if ("isNullObject" in poison) {
          gates.setIsNullObject("Sheet1", poison.isNullObject);
        } else {
          gates.setPoison("Sheet1", poison);
        }
        const result = await adapter.applyWorkbookTemplate(APPLY_DEFAULT);
        expect(result.ok, JSON.stringify({ poison, result })).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
      }
    });

    it("post-write address/rowCount/columnCount mismatch fails after writes", async () => {
      const adapter = new OfficeJsAdapter();
      for (const poison of [{ address: "Sheet1!Z9:Z10" }, { rowCount: 99 }, { columnCount: 77 }]) {
        gates = installTemplateExcel();
        gates.setPostWritePoison("Sheet1", poison);
        const result = await adapter.applyWorkbookTemplate(APPLY_DEFAULT);
        expect(result.ok, JSON.stringify({ poison, result })).toBe(false);
        if (!result.ok) expect(result.unsupported).not.toBe(true);
        expect(gates.writeCalls()).toBeGreaterThan(0);
      }
    });

    it("successful apply has everReadBeforeSync false; premature read is sticky", async () => {
      gates = installTemplateExcel();
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate(APPLY_DEFAULT);
      expect(result.ok).toBe(true);
      expect(gates.everReadBeforeSync()).toBe(false);

      gates = installTemplateExcel();
      gates.forcePrematureRead("Sheet1");
      expect(gates.everReadBeforeSync()).toBe(true);
      // sticky across sync
      await (globalThis as unknown as { Excel: { run: (fn: (c: { sync: () => Promise<void> }) => Promise<void>) => Promise<void> } }).Excel.run(
        async (ctx) => {
          await ctx.sync();
        },
      );
      expect(gates.everReadBeforeSync()).toBe(true);
    });

    it("quoted sheet name with ! applies address/counts", async () => {
      gates = installTemplateExcel({
        extraSheets: [{ name: "A!B", usedAddress: "A1:C2", rows: 2, cols: 3 }],
        activeSheet: "A!B",
      });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        ...APPLY_DEFAULT,
        preset: "financial",
        sheetNames: ["A!B"],
        allSheets: true,
        freezeRows: 0,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.appliedSheets[0]!.name).toBe("A!B");
        expect(result.data.appliedSheets[0]!.range).toBe("A1:C2");
      }
    });

    it("requirement unsupported zero run", async () => {
      const adapter = new OfficeJsAdapter();
      for (const opts of [{ excelApi18: false }, { missingIsSetSupported: true }] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const g = installTemplateExcel(opts);
        const result = await adapter.applyWorkbookTemplate(APPLY_DEFAULT);
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
        expect(g.excelRunCalls()).toBe(0);
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

    it("single Excel.run; fixed syncCount not linear in sheet count", async () => {
      const adapter = new OfficeJsAdapter();
      gates = installTemplateExcel({ sheetCount: 1 });
      const one = await adapter.captureWorkbookTemplate();
      expect(one.ok).toBe(true);
      expect(gates.excelRunCalls()).toBe(1);
      const sync1 = gates.syncCount();
      expect(sync1).toBeGreaterThan(0);
      expect(sync1).toBeLessThanOrEqual(8);
      expect(gates.everReadBeforeSync()).toBe(false);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      gates = installTemplateExcel({ sheetCount: 50 });
      const fifty = await adapter.captureWorkbookTemplate();
      expect(fifty.ok).toBe(true);
      expect(gates.excelRunCalls()).toBe(1);
      const sync50 = gates.syncCount();
      expect(sync50).toBeLessThanOrEqual(8);
      // not linear: 50-sheet syncs should not be ~50x one-sheet
      expect(sync50).toBeLessThanOrEqual(sync1 + 2);
      expect(sync50).toBe(sync1);
      if (fifty.ok) expect(fifty.data.sheetCount).toBe(50);
    });

    it("null mixed succeeds; bad scalars fail (domain)", async () => {
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
        { baseFontName: "" },
        { baseFontName: "   " },
        { baseFontName: 1 },
        { baseFontSize: 0 },
        { baseFontSize: -1 },
        { baseFontSize: NaN },
        { baseFontSize: Infinity },
        { baseFontColor: "red" },
        { baseFontColor: "FFFFFF" },
        { headerFill: true },
        { headerFontColor: 1 },
        { headerBold: "yes" },
        { headerRowHeight: 0 },
        { headerRowHeight: -3 },
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

    it("print scalar poisons ordinary fail", async () => {
      const adapter = new OfficeJsAdapter();
      const printPoisons = [
        { printOrientation: "Port rait" },
        { printOrientation: 1 },
        { printOrientation: "Landscape-Extra" },
        { printPaperSize: "" },
        { printPaperSize: 4 },
        { printFitWide: -1 },
        { printFitWide: 1.5 },
        { printFitWide: "1" },
        { printFitTall: NaN },
        { printAreaIsNull: "false" },
        { printAreaIsNull: false, printAreaAddress: 9 },
        { printHeader: 1 },
        { printFooter: true },
      ] as const;
      for (const poison of printPoisons) {
        gates = installTemplateExcel();
        gates.setPoison("Sheet1", poison as never);
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
      if (!limited.ok) expect(limited.reason).toMatch(/resource-limit/);
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
      }
      expect(gates.everReadBeforeSync()).toBe(false);
    });
  });

  it("WPS typed unsupported for both tools", async () => {
    const wps = new WpsJsaAdapter();
    const apply = await wps.applyWorkbookTemplate({
      ...APPLY_DEFAULT,
      allSheets: true,
    });
    expect(apply.ok).toBe(false);
    if (!apply.ok) expect(apply.unsupported).toBe(true);
    const capture = await wps.captureWorkbookTemplate();
    expect(capture.ok).toBe(false);
    if (!capture.ok) expect(capture.unsupported).toBe(true);
  });
});
