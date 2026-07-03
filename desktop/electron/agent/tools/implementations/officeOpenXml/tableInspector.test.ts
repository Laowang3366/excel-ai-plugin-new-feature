import { mkdtemp, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { inspectOfficeOpenXmlTables } from "./tableInspector";

async function writeZip(filePath: string, files: Record<string, string>): Promise<void> {
  const zip = new JSZip();
  for (const [name, content] of Object.entries(files)) {
    zip.file(name, content);
  }
  await writeFile(filePath, await zip.generateAsync({ type: "nodebuffer" }));
}

describe("tableInspector", () => {
  it("inspects table-like sheet data from an Excel workbook", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-table-test-"));
    try {
      const filePath = path.join(tempDir, "demo.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/worksheets/sheet1.xml": `
          <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr"><is><t>产品</t></is></c>
                <c r="B1" t="inlineStr"><is><t>收入</t></is></c>
              </row>
              <row r="2">
                <c r="A2" t="inlineStr"><is><t>A</t></is></c>
                <c r="B2"><v>120</v></c>
              </row>
            </sheetData>
          </worksheet>
        `,
      });

      const result = await inspectOfficeOpenXmlTables({ filePath });

      expect(result.documentType).toBe("spreadsheet");
      expect(result.tables).toHaveLength(1);
      expect(result.tables[0].columns).toBe(2);
      expect(result.tables[0].rows[0].cells.map((cell) => cell.text)).toEqual(["产品", "收入"]);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("inspects tables from Word document XML", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-table-test-"));
    try {
      const filePath = path.join(tempDir, "demo.docx");
      await writeZip(filePath, {
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

      const result = await inspectOfficeOpenXmlTables({ filePath });

      expect(result.documentType).toBe("word");
      expect(result.tables).toHaveLength(1);
      expect(result.tables[0].rows.length).toBe(2);
      expect(result.tables[0].columns).toBe(2);
      expect(result.tables[0].rows[1].cells[0].text).toBe("小李");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("inspects tables from PowerPoint slide XML", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-table-test-"));
    try {
      const filePath = path.join(tempDir, "demo.pptx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "ppt/slides/slide1.xml": `
          <p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <p:cSld><p:spTree>
              <p:graphicFrame><a:graphic><a:graphicData>
                <a:tbl>
                  <a:tr><a:tc><a:txBody><a:p><a:r><a:t>阶段</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>状态</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                  <a:tr><a:tc><a:txBody><a:p><a:r><a:t>一</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>完成</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
                </a:tbl>
              </a:graphicData></a:graphic></p:graphicFrame>
            </p:spTree></p:cSld>
          </p:sld>
        `,
      });

      const result = await inspectOfficeOpenXmlTables({ filePath, target: "slide:1" });

      expect(result.documentType).toBe("presentation");
      expect(result.tables).toHaveLength(1);
      expect(result.tables[0].rows.length).toBe(2);
      expect(result.tables[0].columns).toBe(2);
      expect(result.tables[0].rows[1].cells[1].text).toBe("完成");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
