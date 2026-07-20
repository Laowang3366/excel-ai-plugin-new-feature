import { describe, expect, it } from "vitest";
import { resolveInitialTabFromSearch } from "../src/resolveInitialTab";

describe("resolveInitialTabFromSearch", () => {
  it("maps whitelist pages", () => {
    expect(resolveInitialTabFromSearch("page=chat")).toBe("chat");
    expect(resolveInitialTabFromSearch("?page=providers")).toBe("providers");
    expect(resolveInitialTabFromSearch(new URLSearchParams("page=host"))).toBe("host");
    expect(resolveInitialTabFromSearch("page=tools")).toBe("tools");
  });

  it("supports legacy settings/model deep-link", () => {
    expect(resolveInitialTabFromSearch("page=settings&section=model")).toBe("providers");
    expect(resolveInitialTabFromSearch("page=settings&section=other")).toBe("chat");
  });

  it("uses first value for duplicate params and ignores case", () => {
    expect(resolveInitialTabFromSearch("page=providers&page=host")).toBe("providers");
    expect(resolveInitialTabFromSearch("page=HOST")).toBe("host");
    expect(resolveInitialTabFromSearch("page=%70roviders")).toBe("providers");
  });

  it("falls back to chat for unknown/malicious/empty input", () => {
    expect(resolveInitialTabFromSearch("")).toBe("chat");
    expect(resolveInitialTabFromSearch(null)).toBe("chat");
    expect(resolveInitialTabFromSearch(undefined)).toBe("chat");
    expect(resolveInitialTabFromSearch("page=admin")).toBe("chat");
    expect(resolveInitialTabFromSearch("page=../../etc/passwd")).toBe("chat");
    expect(resolveInitialTabFromSearch("page=<script>alert(1)</script>")).toBe("chat");
    expect(resolveInitialTabFromSearch("page=chat%00evil")).toBe("chat");
    expect(resolveInitialTabFromSearch("foo=bar")).toBe("chat");
  });
});
