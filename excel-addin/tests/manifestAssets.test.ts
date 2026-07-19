import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import {
  DEFAULT_DEV_BASE_URL,
  DEFAULT_VERSION,
  renderOfficeManifest,
  validateOfficeManifest,
} from "../scripts/officeManifest.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function readPngSize(filePath: string): { width: number; height: number } {
  const buf = readFileSync(filePath);
  expect(buf.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
  const width = buf.readUInt32BE(16);
  const height = buf.readUInt32BE(20);
  return { width, height };
}

describe("Office icons real pixel sizes", () => {
  it("ships 16/32/64/80 PNG with matching IHDR dimensions", () => {
    for (const size of [16, 32, 64, 80] as const) {
      const icon = path.join(root, `public/assets/icon-${size}.png`);
      expect(existsSync(icon), icon).toBe(true);
      const dim = readPngSize(icon);
      expect(dim).toEqual({ width: size, height: size });
    }
  });
});

describe("Office manifest template render/validate", () => {
  const template = readFileSync(
    path.join(root, "manifest/templates/office-excel-manifest.template.xml"),
    "utf8",
  );

  it("renders dev defaults with localhost HTTPS and consistent AppDomain", () => {
    const xml = renderOfficeManifest({
      mode: "dev",
      baseUrl: DEFAULT_DEV_BASE_URL,
      version: DEFAULT_VERSION,
      template,
    });
    const v = validateOfficeManifest(xml, { mode: "dev" });
    expect(v.ok, v.errors.join("; ")).toBe(true);
    expect(xml).toContain("https://localhost:3000/index.html");
    expect(xml).toContain("<AppDomain>https://localhost:3000</AppDomain>");
    expect(xml).toContain("https://localhost:3000/assets/icon-16.png");
    expect(xml).toContain("https://plugin.shelelove.top");
    expect(xml).not.toMatch(/first batch|首批/i);
    expect(xml).toMatch(/Excel AI chat/i);
  });

  it("renders prod with explicit HTTPS base and forbids localhost", () => {
    const xml = renderOfficeManifest({
      mode: "prod",
      baseUrl: "https://plugin.example.com/excel-addin",
      version: "1.2.3.4",
      template,
    });
    const v = validateOfficeManifest(xml, { mode: "prod" });
    expect(v.ok, v.errors.join("; ")).toBe(true);
    expect(xml).toContain("https://plugin.example.com/excel-addin/index.html");
    expect(xml).toContain("<AppDomain>https://plugin.example.com</AppDomain>");
    expect(xml).toContain("<Version>1.2.3.4</Version>");
    expect(xml).not.toMatch(/localhost|127\.0\.0\.1/i);
  });

  it("rejects prod localhost / http / bad version", () => {
    expect(() =>
      renderOfficeManifest({
        mode: "prod",
        baseUrl: "https://localhost:3000",
        template,
      }),
    ).toThrow(/localhost/);
    expect(() =>
      renderOfficeManifest({
        mode: "prod",
        baseUrl: "http://example.com",
        template,
      }),
    ).toThrow(/HTTPS/i);
    expect(() =>
      renderOfficeManifest({
        mode: "dev",
        baseUrl: DEFAULT_DEV_BASE_URL,
        version: "1.0.0",
        template,
      }),
    ).toThrow(/four-part/);
  });

  it("checked-in dev manifest matches template render and validates", () => {
    const checked = readFileSync(
      path.join(root, "manifest/office-excel-manifest.xml"),
      "utf8",
    );
    const expected = renderOfficeManifest({
      mode: "dev",
      baseUrl: DEFAULT_DEV_BASE_URL,
      version: DEFAULT_VERSION,
      template,
    });
    expect(checked.replace(/\r\n/g, "\n").trimEnd()).toBe(
      expected.replace(/\r\n/g, "\n").trimEnd(),
    );
    const v = validateOfficeManifest(checked, { mode: "dev" });
    expect(v.ok, v.errors.join("; ")).toBe(true);
    expect(checked).toMatch(
      /<Id>[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}<\/Id>/i,
    );
    expect(checked).toContain('xsi:type="TaskPaneApp"');
    expect(checked).toContain("<Permissions>ReadWriteDocument</Permissions>");
  });
});
