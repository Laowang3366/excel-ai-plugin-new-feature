import { describe, expect, it } from "vitest";
import {
  decodeHtmlText,
  decodeXmlText,
  escapeXmlAttribute,
  escapeXmlText,
  escapeXmlTextWithQuotes,
} from "./xmlEntities";

describe("xmlEntities", () => {
  it("escapes XML text content without changing quotes", () => {
    expect(escapeXmlText(`A&B <C> "D" 'E'`)).toBe(`A&amp;B &lt;C&gt; "D" 'E'`);
  });

  it("escapes XML attributes and optionally text quotes", () => {
    expect(escapeXmlAttribute(`A&B <C> "D" 'E'`)).toBe(`A&amp;B &lt;C&gt; &quot;D&quot; 'E'`);
    expect(escapeXmlTextWithQuotes(`A&B <C> "D" 'E'`)).toBe(
      "A&amp;B &lt;C&gt; &quot;D&quot; &apos;E&apos;",
    );
  });

  it("decodes XML entities with ampersands last", () => {
    expect(decodeXmlText("&quot;A&quot; &apos;B&apos; &lt;C&gt; &amp; &amp;lt;")).toBe(
      `"A" 'B' <C> & &lt;`,
    );
  });

  it("decodes the HTML entities used by search result parsers", () => {
    expect(decodeHtmlText("A&amp;B &lt;C&gt; &quot;D&quot; &#x27;E&#x27; &#39;F&#39;&nbsp;G")).toBe(
      `A&B <C> "D" 'E' 'F' G`,
    );
  });
});
