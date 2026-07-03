import { beforeEach, describe, expect, test, vi } from "vitest";
import { WordComBridge } from "./wordComBridge";
import { executePowerShell } from "../../../automation/powershell";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

function mockWordConnectedOperation(result: unknown, progId = "Word.Application"): void {
  executePowerShellMock
    .mockResolvedValueOnce(`OK|${progId}|16.0`)
    .mockResolvedValueOnce(JSON.stringify(result));
}

describe("WordComBridge COM lifetime", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("quits the Word app on save when the bridge created it", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({ documentName: "demo.docx", createdApp: true }))
      .mockResolvedValueOnce("");

    const bridge = new WordComBridge();

    await bridge.openDocument("C:\\demo.docx");
    await bridge.saveDocument();

    const saveScript = executePowerShellMock.mock.calls[1][0];
    expect(saveScript).toContain("$shouldQuitApp = $true");
    expect(saveScript).toContain("$doc.Close($false)");
    expect(saveScript).toContain("$app.Quit()");
    expect(saveScript).toContain("[System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)");
    expect(saveScript).toContain("[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)");
  });

  test("clears owned Word app state after a failed save cleanup attempt", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({ documentName: "demo.docx", createdApp: true }))
      .mockRejectedValueOnce(new Error("save failed"))
      .mockResolvedValueOnce("");

    const bridge = new WordComBridge();

    await bridge.openDocument("C:\\demo.docx");
    const failedSave = await bridge.saveDocument();
    await bridge.saveDocument();

    expect(failedSave.success).toBe(false);
    expect(executePowerShellMock.mock.calls[1][0]).toContain("finally");
    expect(executePowerShellMock.mock.calls[2][0]).toContain("$shouldQuitApp = $false");
  });

  test("does not create a blank Word app for active-document operations", async () => {
    mockWordConnectedOperation({});

    const bridge = new WordComBridge();

    await bridge.inspectDocument();

    const inspectScript = executePowerShellMock.mock.calls[1][0];
    expect(inspectScript).not.toContain("New-Object -ComObject");
  });

  test("prefers cached Word ProgID after a successful connection", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({ documentName: "demo.docx", createdApp: false, progId: "Word.Application" }))
      .mockResolvedValueOnce(JSON.stringify({ name: "demo.docx", text: "", charCount: 0, truncated: false }));

    const bridge = new WordComBridge();

    await bridge.openDocument("C:\\demo.docx");
    await bridge.readText();

    const readScript = executePowerShellMock.mock.calls[1][0];
    expect(readScript).toContain("$preferredProgId = 'Word.Application'");
    expect(readScript).toContain("GetActiveObject($preferredProgId)");
    expect(readScript).toContain("foreach ($id in $progIds)");
  });

  test("targets the opened Word document path for later edits", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({
        documentName: "demo.docx",
        fullName: "C:\\docs\\demo.docx",
        createdApp: false,
        progId: "Word.Application",
      }))
      .mockResolvedValueOnce(JSON.stringify({ inserted: true, characters: 3 }));

    const bridge = new WordComBridge();

    await bridge.openDocument("C:\\docs\\demo.docx");
    await bridge.insertText("abc", "end");

    const editScript = executePowerShellMock.mock.calls[1][0];
    expect(editScript).toContain("$_activeDocumentPath = [System.Text.Encoding]::Unicode.GetString");
    expect(editScript).toContain("Resolve-TargetWordDocument");
    expect(editScript).toContain("$doc = Resolve-TargetWordDocument $app $_activeDocumentPath");
    expect(editScript).not.toContain("$doc = $app.ActiveDocument");
  });

  test("validates the target document inside the cached Word operation", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({
        documentName: "demo.docx",
        fullName: "C:\\docs\\demo.docx",
        createdApp: false,
        progId: "Word.Application",
      }))
      .mockResolvedValueOnce(JSON.stringify({ name: "demo.docx", text: "", charCount: 0, truncated: false }));

    const bridge = new WordComBridge();

    await bridge.openDocument("C:\\docs\\demo.docx");
    await bridge.readText();

    expect(executePowerShellMock).toHaveBeenCalledTimes(2);
    expect(executePowerShellMock.mock.calls[1][0]).toContain("GetActiveObject($preferredProgId)");
    expect(executePowerShellMock.mock.calls[1][0]).toContain("Resolve-TargetWordDocument");
    expect(executePowerShellMock.mock.calls[1][0]).toContain("if ($null -eq $doc) { throw '当前没有活动 Word 文档，请先打开或创建文档' }");
  });
});

describe("WordComBridge script parameter injection", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("injects insertHeading level through psVar instead of direct script interpolation", async () => {
    mockWordConnectedOperation({ inserted: true, level: 2 });

    const bridge = new WordComBridge();

    await bridge.insertHeading("阶段总结", 2, "start");

    const headingScript = executePowerShellMock.mock.calls[1][0];
    expect(headingScript).toContain("$_headingLevel = [System.Text.Encoding]::Unicode.GetString");
    expect(headingScript).toContain("$headingLevel = [int]$_headingLevel");
    expect(headingScript).not.toContain("$headingLevel = 2");
  });

  test("keeps empty insertText content as an injected script variable", async () => {
    mockWordConnectedOperation({ inserted: true, characters: 0 });

    const bridge = new WordComBridge();

    await bridge.insertText("", "end");

    const insertScript = executePowerShellMock.mock.calls[1][0];
    expect(insertScript).toContain("$_text = [System.Text.Encoding]::Unicode.GetString");
    expect(insertScript).toContain("characters = $_text.Length");
  });

  test("uses the resolved ProgID variable in inspect output", async () => {
    mockWordConnectedOperation({ name: "demo.docx" });

    const bridge = new WordComBridge();

    await bridge.inspectDocument();

    const inspectScript = executePowerShellMock.mock.calls[1][0];
    expect(inspectScript).toContain("progId = $progId");
    expect(inspectScript).not.toContain("progId = _progId");
  });
});
