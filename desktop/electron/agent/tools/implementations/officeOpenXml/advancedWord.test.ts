import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyWordAdvancedAction } from "./advancedWord";

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

describe("applyWordAdvancedAction", () => {
  it("creates a Word document with title and paragraphs without Python dependencies", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-word-create-"));
    try {
      const filePath = path.join(tempDir, "created.docx");
      const result = await applyWordAdvancedAction({
        operation: "createDocument",
        filePath,
        action: "insert",
        params: {
          title: "测试报告",
          paragraphs: ["第一段", "第二段"],
        },
      });

      const xml = await readZipText(filePath, "word/document.xml");
      const stylesXml = await readZipText(filePath, "word/styles.xml");
      expect(result.status).toBe("done");
      expect(result.engine).toBe("openxml");
      expect(xml).toContain("测试报告");
      expect(xml).toContain("第一段");
      expect(xml).toContain("第二段");
      expect(stylesXml).toContain('w:styleId="Title"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("applies heading style to matching Word paragraphs", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-word-advanced-"));
    try {
      const filePath = path.join(tempDir, "report.docx");
      const outputPath = path.join(tempDir, "report-edited.docx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": "<w:document><w:body><w:p><w:r><w:t>一、概览</w:t></w:r></w:p></w:body></w:document>",
      });

      const result = await applyWordAdvancedAction({
        operation: "applyHeadingStyles",
        filePath,
        outputPath,
        params: { startsWith: "一、", level: 1 },
      });

      const xml = await readZipText(outputPath, "word/document.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('w:val="Heading1"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("replaces non-self-closing paragraph style tags", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-word-advanced-"));
    try {
      const filePath = path.join(tempDir, "styled.docx");
      const outputPath = path.join(tempDir, "styled-edited.docx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": [
          "<w:document><w:body><w:p>",
          '<w:pPr><w:pStyle w:val="OldStyle"></w:pStyle></w:pPr>',
          "<w:r><w:t>二、详情</w:t></w:r>",
          "</w:p></w:body></w:document>",
        ].join(""),
      });

      const result = await applyWordAdvancedAction({
        operation: "applyHeadingStyles",
        filePath,
        outputPath,
        params: { startsWith: "二、", level: 2 },
      });

      const xml = await readZipText(outputPath, "word/document.xml");
      expect(result.status).toBe("done");
      expect(xml).toContain('<w:pStyle w:val="Heading2" />');
      expect(xml).not.toContain("OldStyle");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
