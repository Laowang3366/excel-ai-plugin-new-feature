import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { applyExcelAdvancedAction } from "./advancedExcel";

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

describe("applyExcelAdvancedAction", () => {
  it("creates an Excel workbook with initial data without Python dependencies", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-create-"));
    try {
      const filePath = path.join(tempDir, "created.xlsx");
      const result = await applyExcelAdvancedAction({
        operation: "createWorkbook",
        filePath,
        action: "insert",
        params: {
          sheetNames: ["Data"],
          values: [["A", "B"], [1, 2]],
        },
      });

      const workbookXml = await readZipText(filePath, "xl/workbook.xml");
      const sheetXml = await readZipText(filePath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(result.engine).toBe("openxml");
      expect(workbookXml).toContain('name="Data"');
      expect(sheetXml).toContain('r="A1"');
      expect(sheetXml).toContain("<t>A</t>");
      expect(sheetXml).toContain("<v>2</v>");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("writes a range into an existing workbook through Open XML", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-write-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/workbook.xml": `
          <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>
          </workbook>
        `,
        "xl/_rels/workbook.xml.rels": `
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml" />
          </Relationships>
        `,
        "xl/worksheets/sheet1.xml": '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>keep</t></is></c></row></sheetData></worksheet>',
      });

      const result = await applyExcelAdvancedAction({
        operation: "writeRange",
        filePath,
        target: "range:Sheet1!B2",
        params: { values: [["Name", "Score"], ["Alice", 95]] },
      });

      const sheetXml = await readZipText(filePath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain("<t>keep</t>");
      expect(sheetXml).toContain('r="B2"');
      expect(sheetXml).toContain("<t>Name</t>");
      expect(sheetXml).toContain('r="C3"');
      expect(sheetXml).toContain("<v>95</v>");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("preserves row formatting and non-cell children when writing a range", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-row-format-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/workbook.xml": `
          <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>
          </workbook>
        `,
        "xl/_rels/workbook.xml.rels": `
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml" />
          </Relationships>
        `,
        "xl/worksheets/sheet1.xml": [
          "<worksheet><sheetData>",
          '<row r="2" ht="24" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" s="5" customFormat="1">',
          '<c r="A2" t="inlineStr"><is><t>keep</t></is></c>',
          '<extLst><ext uri="keep"/></extLst>',
          "</row>",
          "</sheetData></worksheet>",
        ].join(""),
      });

      const result = await applyExcelAdvancedAction({
        operation: "writeRange",
        filePath,
        target: "range:Sheet1!B2",
        params: { values: [["new"]] },
      });

      const sheetXml = await readZipText(filePath, "xl/worksheets/sheet1.xml");
      const rowXml = /<row\b[^>]*\br="2"[^>]*>[\s\S]*?<\/row>/.exec(sheetXml)?.[0] || "";
      expect(result.status).toBe("done");
      expect(rowXml).toContain('ht="24"');
      expect(rowXml).toContain('customHeight="1"');
      expect(rowXml).toContain('hidden="1"');
      expect(rowXml).toContain('outlineLevel="2"');
      expect(rowXml).toContain('collapsed="1"');
      expect(rowXml).toContain('s="5"');
      expect(rowXml).toContain('customFormat="1"');
      expect(rowXml).toContain("<t>keep</t>");
      expect(rowXml).toContain('r="B2"');
      expect(rowXml).toContain('<extLst><ext uri="keep"/></extLst>');
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("writes dynamic array formulas with Open XML array metadata and clears spill placeholders", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-dynamic-array-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/workbook.xml": `
          <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>
          </workbook>
        `,
        "xl/_rels/workbook.xml.rels": `
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml" />
          </Relationships>
        `,
        "xl/worksheets/sheet1.xml": [
          "<worksheet><sheetData>",
          '<row r="1"><c r="A1" t="inlineStr"><is><t>keep</t></is></c></row>',
          '<row r="2"><c r="B2" t="inlineStr"><is><t>blocking</t></is></c></row>',
          '<row r="3"><c r="B3" t="inlineStr"><is><t>blocking</t></is></c></row>',
          "</sheetData></worksheet>",
        ].join(""),
      });

      const result = await applyExcelAdvancedAction({
        operation: "writeRange",
        filePath,
        target: "range:Sheet1!B2:B3",
        params: { values: [["=FILTER(A:A,A:A>0)"], [""]] },
      });

      const sheetXml = await readZipText(filePath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain('r="A1"');
      expect(sheetXml).toContain('r="B2"');
      expect(sheetXml).toContain('<f t="array" ref="B2:B3">_xlfn._xlws.FILTER(A:A,A:A&gt;0)</f>');
      expect(sheetXml).not.toContain('r="B3"');
      expect(sheetXml).not.toContain("blocking");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("adds data validation to an Excel worksheet copy", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-advanced-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-edited.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/worksheets/sheet1.xml": "<worksheet><sheetData /></worksheet>",
      });

      const result = await applyExcelAdvancedAction({
        operation: "setDataValidation",
        filePath,
        outputPath,
        target: "range:Sheet1!A2:A10",
        params: { type: "list", values: ["通过", "失败"] },
      });

      const sheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain("<dataValidations");
      expect(sheetXml).toContain("通过,失败");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("preserves existing data validations when adding a new rule", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-validation-merge-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-edited.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/worksheets/sheet1.xml": [
          "<worksheet><sheetData />",
          '<dataValidations count="1" disablePrompts="1">',
          '<dataValidation type="whole" sqref="C2:C10">',
          "<formula1>1</formula1><formula2>10</formula2>",
          "</dataValidation>",
          "</dataValidations>",
          "</worksheet>",
        ].join(""),
      });

      const result = await applyExcelAdvancedAction({
        operation: "setDataValidation",
        filePath,
        outputPath,
        target: "range:Sheet1!A2:A10",
        params: { type: "list", values: ["通过", "失败"] },
      });

      const sheetXml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      expect(result.status).toBe("done");
      expect(sheetXml).toContain('sqref="C2:C10"');
      expect(sheetXml).toContain('sqref="A2:A10"');
      expect(sheetXml).toContain('disablePrompts="1"');
      expect(sheetXml.match(/<dataValidation\b/g)).toHaveLength(2);
      expect(sheetXml).toMatch(/<dataValidations\b[^>]*\bcount="2"/);
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });

  it("resolves range targets to the matching worksheet part", async () => {
    const tempDir = await mkdtemp(path.join(os.tmpdir(), "openxml-excel-advanced-sheet-"));
    try {
      const filePath = path.join(tempDir, "book.xlsx");
      const outputPath = path.join(tempDir, "book-edited.xlsx");
      await writeZip(filePath, {
        "[Content_Types].xml": "<Types />",
        "xl/workbook.xml": `
          <workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
            <sheets>
              <sheet name="Sheet1" sheetId="1" r:id="rId1" />
              <sheet name="Q2" sheetId="2" r:id="rId2" />
            </sheets>
          </workbook>
        `,
        "xl/_rels/workbook.xml.rels": `
          <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml" />
            <Relationship Id="rId2" Type="worksheet" Target="worksheets/sheet2.xml" />
          </Relationships>
        `,
        "xl/worksheets/sheet1.xml": "<worksheet><sheetData /></worksheet>",
        "xl/worksheets/sheet2.xml": "<worksheet><sheetData /></worksheet>",
      });

      const result = await applyExcelAdvancedAction({
        operation: "setDataValidation",
        filePath,
        outputPath,
        target: "range:Q2!B2:B5",
        params: { values: ["确认"] },
      });

      const sheet1Xml = await readZipText(outputPath, "xl/worksheets/sheet1.xml");
      const sheet2Xml = await readZipText(outputPath, "xl/worksheets/sheet2.xml");
      expect(result.status).toBe("done");
      expect(sheet1Xml).not.toContain("<dataValidations");
      expect(sheet2Xml).toContain("<dataValidations");
    } finally {
      await rm(tempDir, { recursive: true, force: true });
    }
  });
});
