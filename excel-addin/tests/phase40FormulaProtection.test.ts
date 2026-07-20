import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { buildArgsPreview } from "../shared/agentChat/approvalPreview";
import { dispositionForRisk } from "../shared/agentChat/approvalPolicy";
import { installFormulaProtectionExcel } from "./fakes/officeJsFormulaProtectionFake";
import { MockHostAdapter } from "./mockHost";

describe("phase40 formula.protection", () => {
  describe("schema + risk", () => {
    it("registers inspect safe and manage dangerous with closed schemas", () => {
      const inspect = TOOL_DEFINITIONS.find((t) => t.name === "formula.protection.inspect");
      const manage = TOOL_DEFINITIONS.find((t) => t.name === "formula.protection.manage");
      expect(inspect?.riskLevel).toBe("safe");
      expect(manage?.riskLevel).toBe("dangerous");
      expect(inspect?.parameters.additionalProperties).toBe(false);
      expect(manage?.parameters.additionalProperties).toBe(false);
      expect(manage?.parameters.properties).toMatchObject({
        command: { enum: ["lock", "unlock"] },
        scope: { enum: ["workbook", "sheet", "target"] },
        password: { type: "string" },
      });
      expect(dispositionForRisk("dangerous")).toBe("approval");
      expect(dispositionForRisk("safe")).toBe("direct");
    });

    it("approval preview redacts password", () => {
      const preview = buildArgsPreview({
        command: "lock",
        scope: "sheet",
        sheetName: "S",
        password: "super-secret",
      }) as Record<string, unknown>;
      expect(preview.password).toBe("[REDACTED]");
      expect(JSON.stringify(preview)).not.toContain("super-secret");
    });
  });

  describe("Office.js", () => {
    let fake: ReturnType<typeof installFormulaProtectionExcel>;
    beforeEach(() => {
      fake = installFormulaProtectionExcel({
        hostSheetName: "HostSheet",
        formulas: [
          ["H1", "H2"],
          ["=A1+1", "input"],
          ["=SUM(A1)", "x"],
        ],
        // mixed lock: first formula locked, second unlocked
        locked: [
          [true, true],
          [true, false],
          [false, false],
        ],
        sheetProtected: false,
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      delete (globalThis as { window?: unknown }).window;
    });

    it("inspect counts formulas and mixed locks without whole-sheet fake", async () => {
      const result = await new OfficeJsAdapter().inspectFormulaProtection({
        scope: "sheet",
        sheetName: "Sheet1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.formulaCount).toBe(2);
        expect(result.data.lockedFormulaCount).toBe(1);
        expect(result.data.sheets[0]?.sheetName).toBe("HostSheet");
        expect(result.data.sheets[0]?.sheetProtected).toBe(false);
        expect(JSON.stringify(result.data)).not.toMatch(/password/i);
      }
    });

    it("inspect empty formulas returns zeros", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installFormulaProtectionExcel({
        formulas: [
          ["a", "b"],
          ["c", "d"],
        ],
        locked: [
          [true, true],
          [true, true],
        ],
      });
      const result = await new OfficeJsAdapter().inspectFormulaProtection({
        scope: "target",
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.formulaCount).toBe(0);
        expect(result.data.lockedFormulaCount).toBe(0);
      }
    });

    it("lock formula cells only, unlockInputs range, protectSheet, verify", async () => {
      const result = await new OfficeJsAdapter().manageFormulaProtection({
        command: "lock",
        scope: "sheet",
        sheetName: "Sheet1",
        password: "pw-never-echo",
        unlockInputs: true,
        protectSheet: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.verified).toBe(true);
        expect(result.data.protection.lockedFormulaCount).toBe(
          result.data.protection.formulaCount,
        );
        expect(result.data.protection.sheets[0]?.sheetProtected).toBe(true);
        expect(JSON.stringify(result.data)).not.toContain("pw-never-echo");
        expect(JSON.stringify(result.data)).not.toMatch(/"password"/);
        expect(result.data.limitations.some((l) => l.includes("unlockInputs"))).toBe(true);
      }
      expect(fake.protectCalls()).toBeGreaterThanOrEqual(1);
      // formula cells (1,0) and (2,0) locked true after
      const matrix = fake.locked();
      expect(matrix[1]![0]).toBe(true);
      expect(matrix[2]![0]).toBe(true);
    });

    it("unlock clears formula locks", async () => {
      await new OfficeJsAdapter().manageFormulaProtection({
        command: "lock",
        scope: "sheet",
        sheetName: "Sheet1",
        protectSheet: false,
        unlockInputs: false,
      });
      const unlocked = await new OfficeJsAdapter().manageFormulaProtection({
        command: "unlock",
        scope: "sheet",
        sheetName: "Sheet1",
        protectSheet: false,
      });
      expect(unlocked.ok).toBe(true);
      if (unlocked.ok) {
        expect(unlocked.data.protection.lockedFormulaCount).toBe(0);
      }
    });

    it("ExcelApi 1.2 false is unsupported without writes", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installFormulaProtectionExcel({ excelApi12: false });
      const result = await new OfficeJsAdapter().manageFormulaProtection({
        command: "lock",
        scope: "sheet",
        sheetName: "Sheet1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
      expect(f.lockWrites()).toHaveLength(0);
    });

    it("missing protection member is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installFormulaProtectionExcel({ hasProtection: false });
      const result = await new OfficeJsAdapter().manageFormulaProtection({
        command: "lock",
        scope: "sheet",
        sheetName: "Sheet1",
        protectSheet: false,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason ?? "").toMatch(/protection\.locked missing/i);
      }
    });

    it("protectSheet verify failure surfaces error without password", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installFormulaProtectionExcel({ failVerifyProtect: true });
      const result = await new OfficeJsAdapter().manageFormulaProtection({
        command: "lock",
        scope: "sheet",
        sheetName: "Sheet1",
        password: "secret-xyz",
        protectSheet: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(JSON.stringify(result)).not.toContain("secret-xyz");
        expect(result.reason ?? "").toMatch(/protectSheet|protected/i);
      }
    });
  });

  describe("executor + mock", () => {
    it("routes tools, rejects unknown fields, strips password from results", async () => {
      const host = new MockHostAdapter();
      host.formulaProtectionSheets.set("Sheet1", {
        protected: false,
        address: "Sheet1!A1:B2",
        formulas: [
          ["=1", "v"],
          ["=2", "w"],
        ],
        locked: [
          [false, true],
          [false, true],
        ],
      });
      const executor = new ToolExecutor(host);

      const inspected = await executor.execute({
        name: "formula.protection.inspect",
        arguments: { scope: "sheet", sheetName: "Sheet1" },
      });
      expect(inspected.ok).toBe(true);
      if (inspected.ok) {
        expect(inspected.data).toMatchObject({ formulaCount: 2, lockedFormulaCount: 0 });
      }

      const managed = await executor.execute({
        name: "formula.protection.manage",
        arguments: {
          command: "lock",
          scope: "sheet",
          sheetName: "Sheet1",
          password: "do-not-leak",
          protectSheet: true,
        },
      });
      expect(managed.ok).toBe(true);
      if (managed.ok) {
        expect(JSON.stringify(managed.data)).not.toContain("do-not-leak");
        expect(managed.data).toMatchObject({ verified: true });
        expect((managed.data as { protection: { lockedFormulaCount: number } }).protection
          .lockedFormulaCount).toBe(2);
      }

      const bad = await executor.execute({
        name: "formula.protection.manage",
        arguments: {
          command: "lock",
          scope: "sheet",
          sheetName: "Sheet1",
          extra: true,
        },
      });
      expect(bad.ok).toBe(false);
      if (!bad.ok) expect(bad.error).toMatch(/unknown field/);

      const missingRange = await executor.execute({
        name: "formula.protection.inspect",
        arguments: { scope: "target", sheetName: "Sheet1" },
      });
      expect(missingRange.ok).toBe(false);
    });
  });

  describe("WPS", () => {
    it("returns typed unsupported", async () => {
      const executor = new ToolExecutor(new WpsJsaAdapter());
      for (const call of [
        {
          name: "formula.protection.inspect" as const,
          arguments: { scope: "sheet", sheetName: "S" },
        },
        {
          name: "formula.protection.manage" as const,
          arguments: { command: "lock", scope: "sheet", sheetName: "S" },
        },
      ]) {
        const result = await executor.execute(call);
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
          expect(result.error ?? "").toMatch(/not verified|unsupported|WPS/i);
        }
      }
    });
  });
});
