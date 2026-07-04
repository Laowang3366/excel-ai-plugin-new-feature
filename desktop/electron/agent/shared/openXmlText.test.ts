import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  extractOpenXmlParagraphTexts,
  extractOpenXmlText,
  extractOpenXmlTextValues,
  readOpenXmlTextParts,
} from "./openXmlText";

describe("openXmlText", () => {
  it("extracts decoded text from exact OpenXML text tags", () => {
    const xml = `<root><w:t>A&amp;B</w:t><w:t>  </w:t><w:t>C&lt;D</w:t></root>`;

    expect(extractOpenXmlTextValues(xml, { tagName: "w:t" })).toEqual(["A&B", "  ", "C<D"]);
    expect(extractOpenXmlTextValues(xml, { tagName: "w:t", includeEmpty: false })).toEqual(["A&B", "C<D"]);
    expect(extractOpenXmlText(xml, { tagName: "w:t", includeEmpty: false })).toBe("A&B\nC<D");
  });

  it("extracts namespace-agnostic text tags for mixed Office XML", () => {
    const xml = `<root><w:t>Word</w:t><a:t>Slide</a:t><t>Sheet</t></root>`;

    expect(extractOpenXmlTextValues(xml, { namespaceAgnostic: true })).toEqual(["Word", "Slide", "Sheet"]);
  });

  it("extracts paragraph text by joining runs and normalizing whitespace", () => {
    const xml = `
      <w:body>
        <w:p><w:r><w:t>Hello</w:t></w:r><w:r><w:t> world</w:t></w:r></w:p>
        <w:p><w:r><w:t>A   B</w:t></w:r></w:p>
      </w:body>
    `;

    expect(extractOpenXmlParagraphTexts(xml)).toEqual(["Hello world", "A B"]);
  });

  it("falls back to text values when paragraph tags are absent", () => {
    const xml = `<root><w:t>One</w:t><w:t>Two</w:t></root>`;

    expect(extractOpenXmlParagraphTexts(xml)).toEqual(["One", "Two"]);
  });

  it("reads matching ZIP parts into text part summaries", async () => {
    const zip = new JSZip();
    zip.file("word/document.xml", `<w:document><w:t>Main</w:t></w:document>`);
    zip.file("word/header1.xml", `<w:hdr><w:t>Header</w:t></w:hdr>`);
    zip.file("word/styles.xml", `<w:styles><w:t>Ignored</w:t></w:styles>`);

    await expect(readOpenXmlTextParts(zip, /^word\/(?:document|header\d+)\.xml$/, { tagName: "w:t" })).resolves.toEqual([
      { partName: "word/document.xml", text: "Main", textLength: 4 },
      { partName: "word/header1.xml", text: "Header", textLength: 6 },
    ]);
  });
});
