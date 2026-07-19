import { existsSync, readFileSync } from "node:fs";
import { inflateSync } from "node:zlib";
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

function readPngMeta(filePath: string): {
  width: number;
  height: number;
  opaque: number;
  colorCount: number;
  hasBrandColor: boolean;
} {
  const buf = readFileSync(filePath);
  expect(buf.subarray(0, 8).toString("hex")).toBe("89504e470d0a1a0a");
  const width = buf.readUInt32BE(16);
  const height = buf.readUInt32BE(20);
  // minimal IDAT inflate for pixel stats
  let offset = 8;
  const chunks: Buffer[] = [];
  while (offset < buf.length) {
    const len = buf.readUInt32BE(offset);
    const type = buf.subarray(offset + 4, offset + 8).toString("ascii");
    const data = buf.subarray(offset + 8, offset + 8 + len);
    offset += 12 + len;
    if (type === "IDAT") chunks.push(data);
    if (type === "IEND") break;
  }
  const raw = inflateSync(Buffer.concat(chunks));
  const colors = new Set<string>();
  let opaque = 0;
  let hasBrandColor = false;
  for (let y = 0; y < height; y += 1) {
    const rowStart = y * (1 + width * 4) + 1;
    for (let x = 0; x < width; x += 1) {
      const i = rowStart + x * 4;
      const r = raw[i]!;
      const g = raw[i + 1]!;
      const b = raw[i + 2]!;
      const a = raw[i + 3]!;
      colors.add(`${r},${g},${b},${a}`);
      if (a > 10) opaque += 1;
      // deep blue or gold brand tones
      if (a > 200 && ((r < 60 && g < 90 && b > 100) || (r > 200 && g > 120 && b < 40))) {
        hasBrandColor = true;
      }
    }
  }
  return { width, height, opaque, colorCount: colors.size, hasBrandColor };
}

describe("Office icons real pixel sizes", () => {
  it("ships 16/32/64/80 PNG with matching IHDR and non-placeholder pixels", () => {
    for (const size of [16, 32, 64, 80] as const) {
      const icon = path.join(root, `public/assets/icon-${size}.png`);
      expect(existsSync(icon), icon).toBe(true);
      const meta = readPngMeta(icon);
      expect(meta.width).toBe(size);
      expect(meta.height).toBe(size);
      expect(meta.opaque).toBeGreaterThan(size * size * 0.4);
      expect(meta.colorCount).toBeGreaterThanOrEqual(3);
      expect(meta.hasBrandColor).toBe(true);
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
    expect(xml).toContain("https://localhost:3000/assets/icon-64.png");
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

  it("rejects unsafe base URLs that would corrupt XML attributes", () => {
    for (const bad of [
      "https://example.com/a&b",
      "https://example.com/a<b",
      'https://example.com/a"b',
      "https://example.com/a'b",
      "https://example.com/a>b",
    ]) {
      expect(() =>
        renderOfficeManifest({ mode: "prod", baseUrl: bad, template }),
      ).toThrow(/unsafe|Invalid|XML/i);
    }
    // percent-encoded ampersand in path is still unsafe once decoded / present as &
    expect(() =>
      renderOfficeManifest({
        mode: "prod",
        baseUrl: "https://example.com/a%26b",
        template,
      }),
    ).toThrow(/unsafe|XML/i);
    // normal subpath still works
    const ok = renderOfficeManifest({
      mode: "prod",
      baseUrl: "https://example.com/excel-addin",
      template,
    });
    expect(validateOfficeManifest(ok, { mode: "prod" }).ok).toBe(true);
  });

  it("validate fails on bare & in attributes and base path escapes", () => {
    const good = renderOfficeManifest({
      mode: "dev",
      baseUrl: DEFAULT_DEV_BASE_URL,
      template,
    });
    const bareAmp = good.replace(
      "https://localhost:3000/index.html",
      "https://localhost:3000/a&b.html",
    );
    const v1 = validateOfficeManifest(bareAmp, { mode: "dev" });
    expect(v1.ok).toBe(false);
    expect(v1.errors.join(" ")).toMatch(/invalid XML attribute entity|not under base/i);

    // different origin for SourceLocation
    const otherOrigin = good.replaceAll(
      "https://localhost:3000/index.html",
      "https://evil.example/index.html",
    );
    const v2 = validateOfficeManifest(otherOrigin, { mode: "dev" });
    expect(v2.ok).toBe(false);

    // missing icon-64
    const no64 = good.replace(/icon-64\.png/g, "icon-missing.png");
    const v3 = validateOfficeManifest(no64, { mode: "dev" });
    expect(v3.ok).toBe(false);
    expect(v3.errors.join(" ")).toMatch(/icon-64/);

    // path escapes base (commands points outside)
    const escaped = good.replace(
      'id="Commands.Url" DefaultValue="https://localhost:3000/index.html"',
      'id="Commands.Url" DefaultValue="https://localhost:3000/../escape.html"',
    );
    // URL parser normalizes .. — craft path not under base by different directory
    const escaped2 = good.replace(
      'id="Commands.Url" DefaultValue="https://localhost:3000/index.html"',
      'id="Commands.Url" DefaultValue="https://localhost:3000/other/index.html"',
    );
    // root base "/" accepts any path — use subpath base to test escape
    const sub = renderOfficeManifest({
      mode: "prod",
      baseUrl: "https://plugin.example.com/excel-addin",
      template,
    });
    const escapedSub = sub.replace(
      'id="Commands.Url" DefaultValue="https://plugin.example.com/excel-addin/index.html"',
      'id="Commands.Url" DefaultValue="https://plugin.example.com/other/index.html"',
    );
    const v4 = validateOfficeManifest(escapedSub, { mode: "prod" });
    expect(v4.ok).toBe(false);
    expect(v4.errors.join(" ")).toMatch(/not under base/);
    void escaped;
    void escaped2;
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
    expect(checked).toContain("icon-64.png");
  });
});
