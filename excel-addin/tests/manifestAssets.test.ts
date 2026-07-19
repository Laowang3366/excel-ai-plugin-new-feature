import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

describe("Office manifest loadability", () => {
  it("uses a valid UUID and local public icons", () => {
    const manifest = readFileSync(
      path.join(root, "manifest/office-excel-manifest.xml"),
      "utf8",
    );
    const id = manifest.match(/<Id>([^<]+)<\/Id>/)?.[1] ?? "";
    expect(id).toMatch(
      /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i,
    );

    for (const size of [16, 32, 64, 80]) {
      const icon = path.join(root, `public/assets/icon-${size}.png`);
      expect(existsSync(icon), icon).toBe(true);
      expect(manifest).toContain(`/assets/icon-${size === 64 ? 64 : size}.png`);
    }

    expect(manifest).toContain("https://localhost:3000/index.html");
    expect(manifest).toContain('xsi:type="TaskPaneApp"');
    expect(manifest).toContain("ReadWriteDocument");
  });
});
