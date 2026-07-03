import { mkdtemp, readFile, rm } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { inspectOfficeOpenXmlFile, replaceOfficeOpenXmlText } from "./officeOpenXmlEngine";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content);
  }
  const buffer = await zip.generateAsync({ type: "nodebuffer" });
  await import("fs/promises").then((fs) => fs.writeFile(filePath, buffer));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing zip part: ${partName}`);
  return part.async("text");
}

describe("officeOpenXmlEngine", () => {
  it("inspects text from Word document XML without Office COM", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-engine-test-"));
    try {
      const filePath = path.join(tempDir, "demo.docx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": `
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p><w:r><w:t>季度报告</w:t></w:r></w:p>
              <w:p><w:r><w:t>旧产品名收入增长</w:t></w:r></w:p>
            </w:body>
          </w:document>
        `,
      });

      const result = await inspectOfficeOpenXmlFile(filePath);

      expect(result.documentType).toBe("word");
      expect(result.textPartCount).toBe(1);
      expect(result.textPreview).toContain("季度报告");
      expect(result.textPreview).toContain("旧产品名收入增长");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("replaces text in PowerPoint slide XML and writes a new file", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-engine-test-"));
    try {
      const sourcePath = path.join(tempDir, "demo.pptx");
      const outputPath = path.join(tempDir, "demo-edited.pptx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": `
          <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>旧产品名发布计划</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
          </p:sld>
        `,
        "ppt/slides/slide2.xml": `
          <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>旧产品名路线图</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
          </p:sld>
        `,
      });

      const result = await replaceOfficeOpenXmlText({
        filePath: sourcePath,
        outputPath,
        findText: "旧产品名",
        replaceText: "新产品名",
      });

      expect(result.documentType).toBe("presentation");
      expect(result.replacements).toBe(2);
      expect(result.outputPath).toBe(outputPath);
      expect(result.changedParts.map((part) => part.partName)).toEqual([
        "ppt/slides/slide1.xml",
        "ppt/slides/slide2.xml",
      ]);
      await expect(readZipText(outputPath, "ppt/slides/slide1.xml")).resolves.toContain("新产品名发布计划");
      await expect(readZipText(outputPath, "ppt/slides/slide2.xml")).resolves.toContain("新产品名路线图");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("does not replace matching text outside Word text nodes", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-engine-test-"));
    try {
      const sourcePath = path.join(tempDir, "demo.docx");
      const outputPath = path.join(tempDir, "demo-edited.docx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": `
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:p data-label="旧产品名"><w:r><w:t>旧产品名收入增长</w:t></w:r></w:p>
            </w:body>
          </w:document>
        `,
      });

      const result = await replaceOfficeOpenXmlText({
        filePath: sourcePath,
        outputPath,
        findText: "旧产品名",
        replaceText: "新产品名",
      });

      const documentXml = await readZipText(outputPath, "word/document.xml");
      expect(result.replacements).toBe(1);
      expect(documentXml).toContain('data-label="旧产品名"');
      expect(documentXml).toContain("<w:t>新产品名收入增长</w:t>");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("inspects and replaces Excel shared strings without Office COM", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-engine-test-"));
    try {
      const sourcePath = path.join(tempDir, "demo.xlsx");
      const outputPath = path.join(tempDir, "demo-edited.xlsx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": "<Types />",
        "xl/sharedStrings.xml": `
          <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <si><t>季度报告</t></si>
            <si data-label="旧产品名"><t>旧产品名收入增长</t></si>
          </sst>
        `,
        "xl/worksheets/sheet1.xml": `
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
              <row r="1"><c r="A1" t="s"><v>0</v></c></row>
              <row r="2"><c r="A2" t="s"><v>1</v></c></row>
            </sheetData>
          </worksheet>
        `,
      });

      const inspectResult = await inspectOfficeOpenXmlFile(sourcePath);
      const replaceResult = await replaceOfficeOpenXmlText({
        filePath: sourcePath,
        outputPath,
        findText: "旧产品名",
        replaceText: "新产品名",
      });

      const sharedStringsXml = await readZipText(outputPath, "xl/sharedStrings.xml");
      expect(inspectResult.documentType).toBe("spreadsheet");
      expect(inspectResult.textPreview).toContain("季度报告");
      expect(inspectResult.textPreview).toContain("旧产品名收入增长");
      expect(replaceResult.documentType).toBe("spreadsheet");
      expect(replaceResult.replacements).toBe(1);
      expect(sharedStringsXml).toContain('data-label="旧产品名"');
      expect(sharedStringsXml).toContain("<t>新产品名收入增长</t>");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
