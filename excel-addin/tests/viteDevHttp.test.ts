import { describe, expect, it } from "vitest";
import { shouldUseViteDevHttp } from "../scripts/viteDevHttp.mjs";

describe("shouldUseViteDevHttp", () => {
  it("defaults to HTTPS (false)", () => {
    expect(shouldUseViteDevHttp({})).toBe(false);
    expect(shouldUseViteDevHttp({ VITE_DEV_HTTP: "0" })).toBe(false);
    expect(shouldUseViteDevHttp({ npm_lifecycle_event: "dev" })).toBe(false);
    expect(shouldUseViteDevHttp({ npm_lifecycle_event: "preview" })).toBe(false);
  });

  it("enables HTTP via VITE_DEV_HTTP=1", () => {
    expect(shouldUseViteDevHttp({ VITE_DEV_HTTP: "1" })).toBe(true);
  });

  it("enables HTTP via npm_lifecycle_event=dev:http (Windows-safe)", () => {
    expect(shouldUseViteDevHttp({ npm_lifecycle_event: "dev:http" })).toBe(true);
    // lifecycle wins even without env prefix
    expect(
      shouldUseViteDevHttp({ npm_lifecycle_event: "dev:http", VITE_DEV_HTTP: undefined }),
    ).toBe(true);
  });

  it("package.json dev:http script is portable (no Unix env prefix)", async () => {
    const { readFileSync } = await import("node:fs");
    const { resolve, dirname } = await import("node:path");
    const { fileURLToPath } = await import("node:url");
    const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
    const pkg = JSON.parse(readFileSync(resolve(root, "package.json"), "utf8")) as {
      scripts: Record<string, string>;
    };
    const script = pkg.scripts["dev:http"];
    expect(script).toBeTruthy();
    // Must not use FOO=1 cmd form (breaks Windows cmd.exe).
    expect(script).not.toMatch(/^[A-Za-z_][A-Za-z0-9_]*=/);
    expect(script).toMatch(/\bvite\b/);
  });
});
