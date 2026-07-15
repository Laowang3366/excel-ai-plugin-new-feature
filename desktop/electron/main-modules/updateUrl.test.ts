import { describe, expect, it } from "vitest";

import { resolveFinalDownloadUrl } from "./updateUrl";

describe("resolveFinalDownloadUrl", () => {
  const requestUrl = new URL("https://plugin.shelelove.top/releases/windows/patch.zip");

  it("falls back to the signed request URL when Electron omits response.url", () => {
    expect(resolveFinalDownloadUrl(requestUrl, "").toString()).toBe(requestUrl.toString());
    expect(resolveFinalDownloadUrl(requestUrl, "   ").toString()).toBe(requestUrl.toString());
  });

  it("uses the final absolute or relative response URL when available", () => {
    expect(
      resolveFinalDownloadUrl(requestUrl, "https://cdn.example.com/patch.zip").toString(),
    ).toBe("https://cdn.example.com/patch.zip");
    expect(resolveFinalDownloadUrl(requestUrl, "mirror/patch.zip").toString()).toBe(
      "https://plugin.shelelove.top/releases/windows/mirror/patch.zip",
    );
  });
});
