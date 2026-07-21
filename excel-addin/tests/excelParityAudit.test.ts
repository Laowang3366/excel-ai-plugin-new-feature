import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const repoRoot = path.resolve(root, "..");

describe("excel parity audit (Phase60)", () => {
  const audit = readFileSync(path.join(root, "docs/excel-parity-audit.md"), "utf8");
  const matrix = readFileSync(path.join(root, "docs/capability-matrix.md"), "utf8");
  const readme = readFileSync(path.join(root, "README.md"), "utf8");
  const changelog = readFileSync(path.join(repoRoot, "CHANGELOG.md"), "utf8");

  it("records narrow c46362f8 device evidence and transient Ribbon conclusion", () => {
    for (const doc of [audit, matrix, readme, changelog]) {
      expect(doc).toContain("Phase60");
      expect(doc).toContain("c46362f8");
    }
    expect(audit).toContain("12.1.0.26885");
    expect(audit).toContain("current=true");
    expect(audit).toContain("drift=[]");
    expect(audit).toMatch(/冷启动|cold start/i);
    expect(audit).toMatch(/瞬态|transient/i);
    expect(audit).toMatch(/not code regression|非.*代码回归|不能.*代码回归/i);
    expect(audit).toContain('address:"G17"');
    expect(audit).toContain("values:[[null]]");
    expect(audit).toMatch(/Do not expand|不得扩大/i);
  });

  it("does not re-open closed G17 / Ribbon as 待复验", () => {
    for (const doc of [audit, matrix, readme]) {
      expect(doc).not.toContain("待 Ribbon 可点后复验");
      expect(doc).not.toContain("host re-verify pending when Ribbon available");
    }
    expect(matrix).not.toContain("**曾真机打开；更新后需冷启动复验**");
  });

  it("says task pane opens and loads UI but layout completeness fails (no full-render claim)", () => {
    // Positive claims must not use obsolete full-render wording
    for (const doc of [audit, matrix, readme, changelog]) {
      expect(doc).not.toContain("task pane full render");
      // Strip meta "must not claim / 不宣称" lines before scanning 完整渲染
      const claimText = doc
        .split("\n")
        .filter((line) => !/Must \*\*not\*\* claim|不\*\*宣称|不宣称|full-render claim/i.test(line))
        .join("\n");
      expect(claimText).not.toContain("完整渲染");
    }
    expect(readme).toMatch(/成功打开并加载 UI/);
    expect(readme).toMatch(/布局完整性未通过|右侧已测裁剪/);
    expect(matrix).toMatch(/成功打开并加载 UI|opens and loads UI/);
    expect(matrix).toMatch(/布局完整性未过|右侧已测裁剪|right-side clip/);
    expect(audit).toMatch(/opens and loads UI/);
    expect(audit).toMatch(/Fails \(measured\)|right-side clip/);
    expect(changelog).toMatch(/成功打开并加载 UI/);
    expect(changelog).toMatch(/布局完整性未通过|右侧已测裁剪/);
    expect(audit).toMatch(/Ribbon|cold start/i);
    expect(audit).toContain("G17");
  });

  it("keeps honesty boundary against full WPS device pass", () => {
    expect(readme).toContain("不得扩大为其它 WPS 工具全部真机通过");
    expect(matrix).toContain("不得扩大为全部 WPS 工具真机通过");
    expect(audit).toMatch(/Do not expand|不得扩大/i);
    expect(audit).toMatch(/member-probe|implemented\*/);
    expect(changelog).toMatch(/不.*宣称全部 WPS 能力真机通过|仍仅 member-probe/);
  });

  it("inventories desktop window tools, excelCapabilities ops, and 98 TOOL_DEFINITIONS", () => {
    expect(TOOL_DEFINITIONS).toHaveLength(98);
    expect(audit).toContain("98");
    expect(audit).toContain("TOOL_DEFINITIONS");
    expect(audit).toContain("excelCapabilities.ts");
    for (const token of [
      "workbook.open",
      "workbook.create",
      "workbook.switch",
      "workbook.save",
      "selection.get",
      "range.read",
      "formula.context",
      "sheet.operation",
      "macro.run",
      "ui.addControl",
    ]) {
      expect(audit).toContain(token);
    }
    for (const op of [
      "createWorkbook",
      "insertChart",
      "createPowerQuery",
      "exportPdf",
      "exportRangeToWord",
      "inspectFormulaDependencies",
      "manageFormulaProtection",
    ]) {
      expect(audit).toContain(op);
    }
    expect(audit).toMatch(/product boundary|Product boundary|产品边界/i);
    expect(audit).toMatch(/WPS typed unsupported|typed unsupported/i);
    expect(audit).toMatch(/Next batch|Recommended next/i);
    expect(audit).toMatch(/provider-chat|Real-device acceptance/i);
    expect(audit).toContain("wpsJsaSelection.ts");
    expect(audit).toContain("wpsJsaUnsupported.ts");
    expect(audit).toContain("wps-remaining-capability-audit.md");
    expect(matrix).toContain("excel-parity-audit.md");
  });

  it("records WPS task-pane CEF viewport vs visible-width clip evidence", () => {
    expect(audit).toMatch(/1428/);
    expect(audit).toMatch(/646/);
    expect(audit).toMatch(/354/);
    expect(audit).toMatch(/CefBrowserWindow|Chrome_RenderWidgetHostHWND/);
    expect(audit).toMatch(/max-width:\s*720|max-width:720/);
    expect(audit).toMatch(/margin:\s*0 auto|margin:0 auto/);
    expect(audit).toContain("styles.css");
    expect(readme).toMatch(/1428|646|任务窗格布局/);
    expect(matrix).toMatch(/1428|646|任务窗格布局/);
    expect(changelog).toMatch(/Phase60\.1|1428/);
  });

  it("prioritizes WPS-only left-align layout fix then e2e; forbids UA and invented members", () => {
    expect(audit).toMatch(/hostKind|data-host|wps-jsa/);
    expect(audit).toMatch(/left-align|左对齐/);
    expect(audit).toMatch(/Playwright/);
    expect(audit).toMatch(/min-width:\s*0|min-width: 0/);
    expect(audit).toMatch(/No trusted WPS host-API code gap|no additional WPS feature/i);
    expect(audit).toMatch(/do \*\*not\*\* invent|Do \*\*not\*\* invent/i);
    expect(audit).toMatch(/User-Agent|UA/);
    expect(audit).toMatch(/no new deps|禁.*新依赖|no new runtime dependencies/i);
    // Must not claim layout already fixed in this doc-only phase
    expect(audit).toMatch(/device retest pending|真机复测|not device-certified|until master retest/i);
    expect(audit).toMatch(/Phase61|520px|data-host/);
    expect(audit).not.toMatch(/layout completeness passed on device|真机布局已通过/);
  });
});
