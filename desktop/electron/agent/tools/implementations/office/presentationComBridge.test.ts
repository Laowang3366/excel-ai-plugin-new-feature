import { beforeEach, describe, expect, test, vi } from "vitest";
import { PresentationComBridge } from "./presentationComBridge";
import { executePowerShell } from "../../../automation/powershell";

vi.mock("../../../automation/powershell", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../../automation/powershell")>();
  return {
    ...actual,
    executePowerShell: vi.fn(),
  };
});

const executePowerShellMock = vi.mocked(executePowerShell);

function mockPresentationConnectedOperation(result: unknown, progId = "PowerPoint.Application"): void {
  executePowerShellMock
    .mockResolvedValueOnce(`OK|${progId}|16.0`)
    .mockResolvedValueOnce(JSON.stringify(result));
}

describe("PresentationComBridge COM lifetime", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("quits the PowerPoint app on save when the bridge created it", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({ presentationName: "demo.pptx", createdApp: true }))
      .mockResolvedValueOnce("");

    const bridge = new PresentationComBridge();

    await bridge.openPresentation("C:\\demo.pptx");
    await bridge.savePresentation();

    const saveScript = executePowerShellMock.mock.calls[1][0];
    expect(saveScript).toContain("$shouldQuitApp = $true");
    expect(saveScript).toContain("$pres.Close()");
    expect(saveScript).toContain("$app.Quit()");
    expect(saveScript).toContain("[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)");
    expect(saveScript).toContain("[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app)");
  });

  test("clears owned PowerPoint app state after a failed save cleanup attempt", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({ presentationName: "demo.pptx", createdApp: true }))
      .mockRejectedValueOnce(new Error("save failed"))
      .mockResolvedValueOnce("");

    const bridge = new PresentationComBridge();

    await bridge.openPresentation("C:\\demo.pptx");
    const failedSave = await bridge.savePresentation();
    await bridge.savePresentation();

    expect(failedSave.success).toBe(false);
    expect(executePowerShellMock.mock.calls[1][0]).toContain("finally");
    expect(executePowerShellMock.mock.calls[2][0]).toContain("$shouldQuitApp = $false");
  });

  test("does not create a blank PowerPoint app for active-presentation operations", async () => {
    mockPresentationConnectedOperation({});

    const bridge = new PresentationComBridge();

    await bridge.inspectPresentation();

    const inspectScript = executePowerShellMock.mock.calls[1][0];
    expect(inspectScript).not.toContain("New-Object -ComObject");
  });

  test("prefers cached PowerPoint ProgID after a successful connection", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({
        presentationName: "demo.pptx",
        createdApp: false,
        progId: "PowerPoint.Application",
      }))
      .mockResolvedValueOnce(JSON.stringify({ name: "demo.pptx", slideCount: 1, slides: [] }));

    const bridge = new PresentationComBridge();

    await bridge.openPresentation("C:\\demo.pptx");
    await bridge.inspectPresentation();

    const inspectScript = executePowerShellMock.mock.calls[1][0];
    expect(inspectScript).toContain("$preferredProgId = 'PowerPoint.Application'");
    expect(inspectScript).toContain("GetActiveObject($preferredProgId)");
    expect(inspectScript).toContain("foreach ($id in $progIds)");
  });

  test("uses the resolved ProgID variable in inspect output", async () => {
    mockPresentationConnectedOperation({ name: "demo.pptx", slideCount: 1, slides: [] });

    const bridge = new PresentationComBridge();

    await bridge.inspectPresentation();

    const inspectScript = executePowerShellMock.mock.calls[1][0];
    expect(inspectScript).toContain("progId = $progId");
    expect(inspectScript).not.toContain("progId = _progId");
  });

  test("checks the shared WPS process name when detecting presentation status", async () => {
    executePowerShellMock
      .mockResolvedValueOnce("WPP")
      .mockResolvedValueOnce("FAIL")
      .mockResolvedValueOnce("OK|12.0|demo.pptx");

    const bridge = new PresentationComBridge();
    const status = await bridge.detectStatus();

    const processScript = executePowerShellMock.mock.calls[0][0];
    const fallbackVerifyScript = executePowerShellMock.mock.calls[2][0];
    expect(processScript).toContain("@('wpp', 'wps')");
    expect(fallbackVerifyScript).toContain("'Kwpp.Application'");
    expect(status).toMatchObject({ connected: true, host: "wpp", presentationName: "demo.pptx" });
  });

  test("targets the opened presentation path for later edits", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({
        presentationName: "demo.pptx",
        fullName: "C:\\docs\\demo.pptx",
        createdApp: false,
        progId: "PowerPoint.Application",
      }))
      .mockResolvedValueOnce(JSON.stringify({ slideIndex: 2, name: "Slide2" }));

    const bridge = new PresentationComBridge();

    await bridge.openPresentation("C:\\docs\\demo.pptx");
    await bridge.addSlide("title", "body");

    const editScript = executePowerShellMock.mock.calls[1][0];
    expect(editScript).toContain("$_activePresentationPath = [System.Text.Encoding]::Unicode.GetString");
    expect(editScript).toContain("Resolve-TargetPresentation");
    expect(editScript).toContain("$pres = Resolve-TargetPresentation $app $_activePresentationPath");
    expect(editScript).not.toContain("$pres = $app.ActivePresentation");
  });

  test("validates the target presentation inside the cached PowerPoint operation", async () => {
    executePowerShellMock
      .mockResolvedValueOnce(JSON.stringify({
        presentationName: "demo.pptx",
        fullName: "C:\\docs\\demo.pptx",
        createdApp: false,
        progId: "PowerPoint.Application",
      }))
      .mockResolvedValueOnce(JSON.stringify({ name: "demo.pptx", slideCount: 1, slides: [] }));

    const bridge = new PresentationComBridge();

    await bridge.openPresentation("C:\\docs\\demo.pptx");
    await bridge.inspectPresentation();

    expect(executePowerShellMock).toHaveBeenCalledTimes(2);
    expect(executePowerShellMock.mock.calls[1][0]).toContain("GetActiveObject($preferredProgId)");
    expect(executePowerShellMock.mock.calls[1][0]).toContain("Resolve-TargetPresentation");
    expect(executePowerShellMock.mock.calls[1][0]).toContain("if ($null -eq $pres) { throw '当前没有活动 PowerPoint 演示文稿，请先打开或创建文档' }");
  });
});

describe("PresentationComBridge slide layout", () => {
  beforeEach(() => {
    executePowerShellMock.mockReset();
  });

  test("prefers custom layouts before falling back to PpSlideLayout enum values", async () => {
    mockPresentationConnectedOperation({ slideIndex: 2, name: "Slide2" });

    const bridge = new PresentationComBridge();

    await bridge.addSlide("标题", "正文", "title_body");

    const addSlideScript = executePowerShellMock.mock.calls[1][0];
    expect(addSlideScript).toContain("Resolve-CustomSlideLayout");
    expect(addSlideScript).toContain("$customLayout = Resolve-CustomSlideLayout $pres $_layout");
    expect(addSlideScript).toContain("$pres.Slides.AddSlide($pres.Slides.Count + 1, $customLayout)");
    expect(addSlideScript).toContain("$pres.Slides.Add($pres.Slides.Count + 1, $layoutValue)");
  });

  test("keeps empty slide title and body as injected script variables", async () => {
    mockPresentationConnectedOperation({ slideIndex: 2, name: "Slide2" });

    const bridge = new PresentationComBridge();

    await bridge.addSlide("", "", "blank");

    const addSlideScript = executePowerShellMock.mock.calls[1][0];
    expect(addSlideScript).toContain("$_title = [System.Text.Encoding]::Unicode.GetString");
    expect(addSlideScript).toContain("$_body = [System.Text.Encoding]::Unicode.GetString");
    expect(addSlideScript).toContain("if ($_title)");
    expect(addSlideScript).toContain("if ($_body)");
  });
});
