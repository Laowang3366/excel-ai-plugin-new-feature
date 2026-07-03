import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyPresentationAdvancedAction } from "./advancedPresentation";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) zip.file(name, content);
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing ${partName}`);
  return part.async("text");
}

describe("applyPresentationAdvancedAction", () => {
  it("creates a basic presentation package without PowerPoint COM", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-ppt-create-"));
    try {
      const filePath = path.join(tempDir, "health-talk.pptx");

      const result = await applyPresentationAdvancedAction({
        operation: "createPresentation",
        action: "insert",
        filePath,
        params: {
          title: "健康饮食营养讲座",
          subtitle: "均衡膳食与日常习惯",
        },
      });

      const zip = await JSZip.loadAsync(await readFile(filePath));
      expect(result.status).toBe("done");
      expect(zip.file("ppt/presentation.xml")).toBeTruthy();
      expect(zip.file("ppt/slides/slide1.xml")).toBeTruthy();
      expect(await readZipText(filePath, "ppt/slides/slide1.xml")).toContain("健康饮食营养讲座");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("applies theme colors to slide text runs", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-ppt-advanced-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      const outputPath = path.join(tempDir, "slides-edited.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": '<p:sld><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>标题</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      });

      const result = await applyPresentationAdvancedAction({
        operation: "applyTheme",
        filePath,
        outputPath,
        params: { accentColor: "1F4E79" },
      });

      const xml = await readZipText(outputPath, "ppt/slides/slide1.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('val="1F4E79"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("deletes selected slides and removes their presentation relationships", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-ppt-delete-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      const outputPath = path.join(tempDir, "slides-deleted.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/presentation.xml": '<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst></p:presentation>',
        "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="rId1" Target="slides/slide1.xml"/><Relationship Id="rId2" Target="slides/slide2.xml"/><Relationship Id="rId3" Target="slides/slide3.xml"/></Relationships>',
        "ppt/slides/slide1.xml": "<p:sld>one</p:sld>",
        "ppt/slides/slide2.xml": "<p:sld>two</p:sld>",
        "ppt/slides/slide3.xml": "<p:sld>three</p:sld>",
      });

      const result = await applyPresentationAdvancedAction({
        operation: "deleteSlides",
        action: "edit",
        filePath,
        outputPath,
        params: { slides: [2, 3] },
      });

      const zip = await JSZip.loadAsync(await readFile(outputPath));
      const presentationXml = await readZipText(outputPath, "ppt/presentation.xml");
      const relsXml = await readZipText(outputPath, "ppt/_rels/presentation.xml.rels");
      expect(result.status).toBe("done");
      expect(result.changes).toHaveLength(2);
      expect(presentationXml).toContain('r:id="rId1"');
      expect(presentationXml).not.toContain('r:id="rId2"');
      expect(presentationXml).not.toContain('r:id="rId3"');
      expect(relsXml).toContain('Target="slides/slide1.xml"');
      expect(relsXml).not.toContain('Target="slides/slide2.xml"');
      expect(relsXml).not.toContain('Target="slides/slide3.xml"');
      expect(zip.file("ppt/slides/slide1.xml")).toBeTruthy();
      expect(zip.file("ppt/slides/slide2.xml")).toBeNull();
      expect(zip.file("ppt/slides/slide3.xml")).toBeNull();
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("adds content slides to an existing presentation through Open XML", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-ppt-add-slides-"));
    try {
      const filePath = path.join(tempDir, "slides.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": [
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
          '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>',
          "</Types>",
        ].join(""),
        "ppt/presentation.xml": [
          '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>',
          '<p:sldSz cx="12192000" cy="6858000" type="wide"/>',
          "</p:presentation>",
        ].join(""),
        "ppt/_rels/presentation.xml.rels": [
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>',
          "</Relationships>",
        ].join(""),
        "ppt/slides/slide1.xml": "<p:sld>one</p:sld>",
      });

      const result = await applyPresentationAdvancedAction({
        operation: "addSlides",
        action: "insert",
        filePath,
        params: {
          slides: [
            { title: "营养原则", bullets: ["均衡膳食", "适量运动"] },
            { title: "每日建议", body: "早餐\n午餐\n晚餐" },
          ],
        },
      });

      const zip = await JSZip.loadAsync(await readFile(filePath));
      const presentationXml = await readZipText(filePath, "ppt/presentation.xml");
      const relsXml = await readZipText(filePath, "ppt/_rels/presentation.xml.rels");
      const contentTypesXml = await readZipText(filePath, "[Content_Types].xml");
      const slide2Xml = await readZipText(filePath, "ppt/slides/slide2.xml");
      const slide3Xml = await readZipText(filePath, "ppt/slides/slide3.xml");
      expect(result.status).toBe("done");
      expect(zip.file("ppt/slides/_rels/slide2.xml.rels")).toBeTruthy();
      expect(zip.file("ppt/slides/_rels/slide3.xml.rels")).toBeTruthy();
      expect(presentationXml).toContain('id="257" r:id="rId2"');
      expect(presentationXml).toContain('id="258" r:id="rId3"');
      expect(relsXml).toContain('Target="slides/slide2.xml"');
      expect(relsXml).toContain('Target="slides/slide3.xml"');
      expect(contentTypesXml).toContain('PartName="/ppt/slides/slide2.xml"');
      expect(contentTypesXml).toContain('PartName="/ppt/slides/slide3.xml"');
      expect(slide2Xml).toContain("营养原则");
      expect(slide2Xml).toContain("均衡膳食");
      expect(slide3Xml).toContain("每日建议");
      expect(slide3Xml).toContain("早餐");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
