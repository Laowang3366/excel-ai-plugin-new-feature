import { mkdtemp, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { inspectOfficeOpenXmlLayout } from "./layoutInspector";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content);
  }
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

describe("layoutInspector", () => {
  it("inspects text objects from a PowerPoint slide", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-layout-test-"));
    try {
      const filePath = path.join(tempDir, "demo.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": `
          <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <p:cSld>
              <p:spTree>
                <p:sp><p:txBody><a:p><a:r><a:t>销售复盘</a:t></a:r></a:p></p:txBody></p:sp>
                <p:pic><p:blipFill><a:blip r:embed="rId2" /></p:blipFill></p:pic>
              </p:spTree>
            </p:cSld>
          </p:sld>
        `,
      });

      const result = await inspectOfficeOpenXmlLayout({ filePath, target: "slide:1" });

      expect(result.documentType).toBe("presentation");
      expect(result.objectCount).toBe(1);
      expect(result.objects[0]).toMatchObject({
        type: "text",
        partName: "ppt/slides/slide1.xml",
        text: "销售复盘",
        textLength: 4,
      });
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
