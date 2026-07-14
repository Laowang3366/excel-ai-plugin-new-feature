import { describe, expect, it } from "vitest";

import { isAllowedWindowNavigation } from "./windowNavigationPolicy";

describe("window navigation policy", () => {
  it("allows only the configured application document", () => {
    const appUrl = "file:///C:/Program%20Files/Wengge/dist/index.html";
    expect(isAllowedWindowNavigation(`${appUrl}#settings`, appUrl)).toBe(true);
    expect(isAllowedWindowNavigation("file:///C:/Windows/System32/drivers/etc/hosts", appUrl)).toBe(false);
    expect(isAllowedWindowNavigation("https://attacker.example", appUrl)).toBe(false);
  });

  it("does not trust another host that resembles the development origin", () => {
    const appUrl = "http://localhost:5173/";
    expect(isAllowedWindowNavigation("http://localhost:5173/#chat", appUrl)).toBe(true);
    expect(isAllowedWindowNavigation("http://localhost.attacker.example:5173/", appUrl)).toBe(false);
  });
});
