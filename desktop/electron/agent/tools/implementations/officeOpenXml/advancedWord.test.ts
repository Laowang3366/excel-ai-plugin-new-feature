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

  it.each(["header", "footer"] as const)(
    "connects a %s part to the document relationship graph",
    async (kind) => {
      const tempDir = await mkdtemp(path.join(os.tmpdir(), `openxml-word-${kind}-`));
      try {
        const filePath = path.join(tempDir, "report.docx");
        const outputPath = path.join(tempDir, `report-${kind}.docx`);
        await applyWordAdvancedAction({
          operation: "createDocument",
          filePath,
          action: "insert",
          params: { paragraphs: ["正文"] },
        });

        const result = await applyWordAdvancedAction({
          operation: "setHeaderFooter",
          filePath,
          outputPath,
          params: { kind, text: `${kind}-text` },
        });

        const partName = kind === "footer" ? "word/footer1.xml" : "word/header1.xml";
        const relationshipType = kind === "footer" ? "footer" : "header";
        const referenceName = kind === "footer" ? "footerReference" : "headerReference";
        const contentType = kind === "footer"
          ? "wordprocessingml.footer+xml"
          : "wordprocessingml.header+xml";
        const partXml = await readZipText(outputPath, partName);
        const documentXml = await readZipText(outputPath, "word/document.xml");
        const relsXml = await readZipText(outputPath, "word/_rels/document.xml.rels");
        const contentTypesXml = await readZipText(outputPath, "[Content_Types].xml");
        const relationship = new RegExp(
          `<Relationship\\b[^>]*Id="([^"]+)"[^>]*Type="[^"]+/relationships/${relationshipType}"[^>]*Target="([^"]+)"`
        ).exec(relsXml);

        expect(result.status).toBe("done");
        expect(partXml).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
        expect(partXml).toContain(`${kind}-text`);
        expect(relationship).not.toBeNull();
        expect(relationship?.[2]).toBe(`${relationshipType}1.xml`);
        expect(documentXml).toContain('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"');
        expect(documentXml).toContain(
          `w:${referenceName} w:type="default" r:id="${relationship?.[1]}"`
        );
        expect(contentTypesXml).toContain(`PartName="/word/${relationshipType}1.xml"`);
        expect(contentTypesXml).toContain(contentType);
        expect(result.changes.map((change) => change.target)).toEqual(expect.arrayContaining([
          partName,
          "word/document.xml",
          "word/_rels/document.xml.rels",
          "[Content_Types].xml",
        ]));
        expect(result.validation?.checks.map((check) => check.name)).toEqual(expect.arrayContaining([
          "header-footer-part",
          "header-footer-relationship",
          "section-reference",
          "content-type",
        ]));
      } finally {
        await rm(tempDir, { recursive: true, force: true });
      }
    }
  );

  it("keeps header references before footer references when both are set", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-word-header-footer-"));
    try {
      const filePath = path.join(tempDir, "report.docx");
      const headerPath = path.join(tempDir, "report-header.docx");
      const outputPath = path.join(tempDir, "report-header-footer.docx");
      await applyWordAdvancedAction({
        operation: "createDocument",
        filePath,
        action: "insert",
        params: { paragraphs: ["正文"] },
      });

      await applyWordAdvancedAction({
        operation: "setHeaderFooter",
        filePath,
        outputPath: headerPath,
        params: { kind: "header", text: "页眉" },
      });
      const result = await applyWordAdvancedAction({
        operation: "setHeaderFooter",
        filePath: headerPath,
        outputPath,
        params: { kind: "footer", text: "页脚" },
      });

      const documentXml = await readZipText(outputPath, "word/document.xml");
      const sectionXml = /<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/.exec(documentXml)?.[0] || "";
      const headerIndex = sectionXml.indexOf("<w:headerReference");
      const footerIndex = sectionXml.indexOf("<w:footerReference");

      expect(result.status).toBe("done");
      expect(headerIndex).toBeGreaterThanOrEqual(0);
      expect(footerIndex).toBeGreaterThan(headerIndex);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
