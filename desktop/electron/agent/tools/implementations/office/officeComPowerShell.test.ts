import { beforeEach, describe, expect, test, vi } from "vitest";
import {
  buildAcquireOfficeAppScript,
  buildTargetOfficeFileResolverScript,
  findActiveOfficeComProgId,
  psNullableVar,
  verifyDirectOfficeCom,
  verifyOfficeComAvailable,
} from "./officeComPowerShell";
import { executePowerShell } from "../../../automation/powershell";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

describe("officeComPowerShell", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("builds acquire scripts with preferred ProgID and app-specific errors", () => {
    const script = buildAcquireOfficeAppScript({
      progIds: ["Word.Application", "Kwps.Application"],
      allowCreate: false,
      preferredProgId: "Word.Application",
      missingMessage: "Word is not running",
    });

    expect(script).toContain("$progIds = @('Word.Application', 'Kwps.Application')");
    expect(script).toContain("$preferredProgId = 'Word.Application'");
    expect(script).toContain("GetActiveObject($preferredProgId)");
    expect(script).not.toContain("New-Object -ComObject");
    expect(script).toContain("throw 'Word is not running'");
  });

  test("builds target file resolver scripts for different Office collections", () => {
    const script = buildTargetOfficeFileResolverScript({
      functionName: "Resolve-TargetPresentation",
      collectionProperty: "Presentations",
      activeProperty: "ActivePresentation",
    });

    expect(script).toContain("function Resolve-TargetPresentation($app, $targetPath)");
    expect(script).toContain("foreach ($candidate in $app.Presentations)");
    expect(script).toContain("$app.ActivePresentation");
  });

  test("injects nullable PowerShell variables consistently", () => {
    expect(psNullableVar("_path", null)).toBe("$_path = $null");
    expect(psNullableVar("_path", "C:\\demo.docx")).toContain("$_path = [System.Text.Encoding]::Unicode.GetString");
  });

  test("verifies COM hosts by host-specific ProgID order", async () => {
    executePowerShellMock.mockResolvedValueOnce("FAIL").mockResolvedValueOnce("OK|16.0|demo.pptx");

    const result = await verifyOfficeComAvailable({
      hosts: ["wpp"],
      defaultHost: "powerpoint",
      progIdsForHost: (host) => host === "wpp" ? ["Wpp.Application", "Kwpp.Application"] : ["PowerPoint.Application"],
      activeObjectExpression: "$app.ActivePresentation",
    });

    expect(result).toMatchObject({
      available: true,
      host: "wpp",
      version: "16.0",
      activeName: "demo.pptx",
      progId: "Kwpp.Application",
    });
  });

  test("verifies direct COM ProgIDs and resolves host from ProgID", async () => {
    executePowerShellMock.mockResolvedValueOnce("FAIL").mockResolvedValueOnce("OK|12.0|demo.docx");

    const result = await verifyDirectOfficeCom({
      progIds: ["Word.Application", "Kwps.Application"],
      defaultHost: "word",
      hostForProgId: (progId) => progId.includes("wps") || progId.includes("Kwps") ? "wps" : "word",
      activeObjectExpression: "$app.ActiveDocument",
    });

    expect(result).toMatchObject({
      available: true,
      host: "wps",
      version: "12.0",
      activeName: "demo.docx",
      progId: "Kwps.Application",
    });
  });

  test("finds the first active COM ProgID for ensureConnected", async () => {
    executePowerShellMock.mockResolvedValueOnce("FAIL").mockResolvedValueOnce("OK|Kwps.Application|12.0");

    const result = await findActiveOfficeComProgId({
      progIds: ["Word.Application", "Kwps.Application"],
      hostForProgId: (progId) => progId.includes("wps") || progId.includes("Kwps") ? "wps" : "word",
    });

    expect(result).toEqual({ progId: "Kwps.Application", host: "wps", version: "12.0" });
  });
});
