import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyOfficeOpenXmlTableStyle } from "./tableStyler";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content);
  }
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

async function readZipText(filePath: string, partName: string): Promise<string> {
  const zip = await JSZip.loadAsync(await readFile(filePath));
  const part = zip.file(partName);
  if (!part) throw new Error(`missing zip part: ${partName}`);
  return part.async("text");
}

describe("tableStyler", () => {
  it("applies professional style to Word table header in an output copy", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-style-test-"));
    try {
      const sourcePath = path.join(tempDir, "demo.docx");
      const outputPath = path.join(tempDir, "demo-styled.docx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": "<Types />",
        "word/document.xml": `
          <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:body>
              <w:tbl>
                <w:tr><w:tc><w:p><w:r><w:t>姓名</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>分数</w:t></w:r></w:p></w:tc></w:tr>
                <w:tr><w:tc><w:p><w:r><w:t>小李</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>95</w:t></w:r></w:p></w:tc></w:tr>
              </w:tbl>
            </w:body>
          </w:document>
        `,
      });

      const sourceBefore = await readZipText(sourcePath, "word/document.xml");
      const result = await applyOfficeOpenXmlTableStyle({
        filePath: sourcePath,
        outputPath,
        style: "professional",
      });
      const sourceAfter = await readZipText(sourcePath, "word/document.xml");
      const outputXml = await readZipText(outputPath, "word/document.xml");

      expect(result.outputPath).toBe(outputPath);
      expect(result.changedParts).toEqual(["word/document.xml"]);
      expect(sourceAfter).toBe(sourceBefore);
      expect(outputXml).toContain('w:fill="1F4E79"');
      expect(outputXml).toContain("<w:b");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("applies professional style to Excel header cells in an output copy", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-style-test-"));
    try {
      const sourcePath = path.join(tempDir, "demo.xlsx");
      const outputPath = path.join(tempDir, "demo-styled.xlsx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": "<Types></Types>",
        "xl/workbook.xml": "<workbook />",
        "xl/_rels/workbook.xml.rels": "<Relationships></Relationships>",
        "xl/worksheets/sheet1.xml": `
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
              <row r="1"><c r="A1" t="inlineStr"><is><t>产品</t></is></c><c r="B1" t="inlineStr"><is><t>收入</t></is></c></row>
              <row r="2"><c r="A2" t="inlineStr"><is><t>A</t></is></c><c r="B2"><v>120</v></c></row>
            </sheetData>
          </worksheet>
        `,
      });

      const result = await applyOfficeOpenXmlTableStyle({
        filePath: sourcePath,
        outputPath,
        style: "professional",
      });
      const outputSheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      const stylesXml = await readZipText(outputPath, "xl/styles.xml");

      expect(result.changedParts).toContain("xl/worksheets/sheet1.xml");
      expect(result.changedParts).toContain("xl/styles.xml");
      expect(outputSheetXml).toContain('r="A1" s="1"');
      expect(stylesXml).toContain('rgb="FF1F4E79"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("preserves existing Excel styles and appends the header style", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-style-preserve-"));
    try {
      const sourcePath = path.join(tempDir, "styled.xlsx");
      const outputPath = path.join(tempDir, "styled-output.xlsx");
      await writeZip(sourcePath, {
        "[Content_Types].xml": [
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />',
          "</Types>",
        ].join(""),
        "xl/workbook.xml": "<workbook />",
        "xl/_rels/workbook.xml.rels": [
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />',
          "</Relationships>",
        ].join(""),
        "xl/styles.xml": [
          '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd" /></numFmts>',
          '<fonts count="2"><font><name val="OriginalFont" /></font><font><b /></font></fonts>',
          '<fills count="2"><fill><patternFill patternType="none" /></fill><fill><patternFill patternType="gray125" /></fill></fills>',
          '<borders count="2"><border><left /><right /><top /><bottom /><diagonal /></border><border><left style="thin"><color rgb="FF000000" /></left><right /><top /><bottom /><diagonal /></border></borders>',
          '<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="1" /></cellStyleXfs>',
          '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="1" /><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" /></cellXfs>',
          '<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0" /><cellStyle name="Custom" xfId="1" /></cellStyles>',
          '<dxfs count="1"><dxf><font><i /></font></dxf></dxfs>',
          "</styleSheet>",
        ].join(""),
        "xl/worksheets/sheet1.xml": [
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',
          '<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>日期</t></is></c><c r="B1" s="1" t="inlineStr"><is><t>金额</t></is></c></row>',
          '<row r="2"><c r="A2" s="2"><v>45800</v></c><c r="B2" s="2"><v>120</v></c></row>',
          "</sheetData></worksheet>",
        ].join(""),
      });

      await applyOfficeOpenXmlTableStyle({
        filePath: sourcePath,
        outputPath,
        style: "professional",
      });

      const outputSheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      const stylesXml = await readZipText(outputPath, "xl/styles.xml");

      expect(stylesXml).toContain('formatCode="yyyy-mm-dd"');
      expect(stylesXml).toContain('<name val="OriginalFont" />');
      expect(stylesXml).toContain('<left style="thin">');
      expect(stylesXml).toContain('<cellStyle name="Custom" xfId="1" />');
      expect(stylesXml).toContain("<dxfs count=\"1\">");
      expect(stylesXml).toContain('<fonts count="3">');
      expect(stylesXml).toContain('<fills count="3">');
      expect(stylesXml).toContain('<cellXfs count="4">');
      expect(stylesXml).toContain('fontId="2" fillId="2"');
      expect(outputSheetXml).toContain('r="A1" s="3"');
      expect(outputSheetXml).toContain('r="A2" s="2"');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
