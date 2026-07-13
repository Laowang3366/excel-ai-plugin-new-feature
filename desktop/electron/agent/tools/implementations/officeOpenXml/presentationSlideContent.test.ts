import { describe, expect, it } from "vitest";
import { contentSlideXml, emptySlideRelsXml, normalizeSlidesParam } from "./presentationSlideContent";

describe("presentationSlideContent", () => {
  it("normalizes slide arrays and bullet lists into body text", () => {
    const slides = normalizeSlidesParam({
      slides: [
        { title: "Agenda", bullets: ["Market", "Roadmap"] },
        { heading: "Risks", content: "Dependencies" },
        "Appendix",
        null,
      ],
    });

    expect(slides).toEqual([
      { title: "Agenda", body: "Market\nRoadmap", layout: undefined },
      { title: "Risks", body: "Dependencies", layout: undefined },
      { title: "Appendix", body: "" },
    ]);
  });

  it("normalizes single-slide params from title and body aliases", () => {
    expect(normalizeSlidesParam({ heading: "Summary", points: ["A", "B"], layout: "blank" })).toEqual([
      { title: "Summary", body: "A\nB", layout: "blank" },
    ]);
    expect(normalizeSlidesParam({})).toEqual([]);
  });

  it("builds content slide xml with escaped text and relationship xml", () => {
    const xml = contentSlideXml("Q&A <2026>", "Use \"quotes\" & apostrophes", undefined);

    expect(xml).toContain("Q&amp;A &lt;2026&gt;");
    expect(xml).toContain("Use &quot;quotes&quot; &amp; apostrophes");
    expect(emptySlideRelsXml()).toContain("<Relationships");
    expect(emptySlideRelsXml()).toContain("relationships/slideLayout");
  });

  it("omits placeholder shapes for a blank empty slide", () => {
    const xml = contentSlideXml("", "", "blank");

    expect(xml).not.toContain("<p:txBody>");
  });
});
