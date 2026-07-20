import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import {
  assertIndexAssetsUnderBase,
  assertNoSensitiveDistPaths,
  assertViteBaseMatchesBaseUrl,
  buildBuildInfo,
  buildSha256Sums,
  deriveViteBaseFromBaseUrl,
  expandPackageVersion,
  makeArtifactName,
  requireFourPartVersion,
  resolvePackageInputs,
} from "../scripts/packageProdCore.mjs";
import {
  renderOfficeManifest,
  validateOfficeManifest,
} from "../scripts/officeManifest.mjs";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

describe("packageProdCore version / base resolution", () => {
  it("expands three-part package version to four-part", () => {
    expect(expandPackageVersion("0.1.0")).toBe("0.1.0.0");
    expect(expandPackageVersion("1.2.3.4")).toBe("1.2.3.4");
    expect(() => expandPackageVersion("1.2")).toThrow(/version/);
    expect(requireFourPartVersion("9.8.7.6")).toBe("9.8.7.6");
    expect(() => requireFourPartVersion("9.8.7")).toThrow(/four-part/);
  });

  it("derives vite base from base_url pathname", () => {
    expect(deriveViteBaseFromBaseUrl("https://example.com/excel-addin")).toBe(
      "/excel-addin/",
    );
    expect(deriveViteBaseFromBaseUrl("https://example.com/")).toBe("/");
    expect(deriveViteBaseFromBaseUrl("https://example.com/a/b")).toBe("/a/b/");
  });

  it("accepts matching explicit vite_base and rejects mismatch", () => {
    expect(
      assertViteBaseMatchesBaseUrl("/excel-addin/", "https://example.com/excel-addin"),
    ).toBe("/excel-addin/");
    expect(
      assertViteBaseMatchesBaseUrl("excel-addin", "https://example.com/excel-addin"),
    ).toBe("/excel-addin/");
    expect(() =>
      assertViteBaseMatchesBaseUrl("/other/", "https://example.com/excel-addin"),
    ).toThrow(/does not match/);
  });

  it("resolvePackageInputs fails on localhost and bad version", () => {
    expect(() =>
      resolvePackageInputs({
        baseUrl: "https://localhost:3000",
        packageJsonVersion: "0.1.0",
      }),
    ).toThrow(/localhost/);
    const ok = resolvePackageInputs({
      baseUrl: "https://example.com/excel-addin",
      packageJsonVersion: "0.1.0",
    });
    expect(ok).toEqual({
      baseUrl: "https://example.com/excel-addin",
      viteBase: "/excel-addin/",
      version: "0.1.0.0",
      packageJsonVersion: "0.1.0",
    });
    expect(() =>
      resolvePackageInputs({
        baseUrl: "https://example.com/excel-addin",
        version: "1.0.0",
        packageJsonVersion: "0.1.0",
      }),
    ).toThrow(/four-part/);
  });
});

describe("packageProdCore dist checks", () => {
  it("asserts local index assets under vite base; ignores CDN", () => {
    const html = `<!doctype html><html><head>
      <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
      <script type="module" src="/excel-addin/assets/app.js"></script>
      <link rel="stylesheet" href="/excel-addin/assets/app.css">
    </head></html>`;
    const refs = assertIndexAssetsUnderBase(html, "/excel-addin/");
    expect(refs).toEqual([
      "assets/app.js",
      "assets/app.css",
    ]);
    expect(() =>
      assertIndexAssetsUnderBase(
        `<script src="/other/assets/x.js"></script>`,
        "/excel-addin/",
      ),
    ).toThrow(/not under VITE_BASE/);
  });

  it("rejects sensitive dist paths", () => {
    expect(() => assertNoSensitiveDistPaths(["assets/a.js", ".env"])).toThrow(
      /sensitive/,
    );
    expect(() =>
      assertNoSensitiveDistPaths(["nested/node_modules/x"]),
    ).toThrow(/sensitive/);
    expect(() => assertNoSensitiveDistPaths(["secret.pem"])).toThrow(/sensitive/);
    expect(() => assertNoSensitiveDistPaths(["CLAUDE.md"])).toThrow(/sensitive/);
    expect(() =>
      assertNoSensitiveDistPaths(["index.html", "assets/a.js"]),
    ).not.toThrow();
    expect(() =>
      assertNoSensitiveDistPaths(["assets/a..b.js"]),
    ).not.toThrow();
  });

  it("builds stable sorted SHA256SUMS and BUILD_INFO without secrets", () => {
    const files = [
      { relativePath: "b.txt", content: "b" },
      { relativePath: "a.txt", content: "a" },
      { relativePath: "SHA256SUMS.txt", content: "ignore" },
    ];
    const sums = buildSha256Sums(files);
    const lines = sums.trim().split("\n");
    expect(lines).toHaveLength(2);
    expect(lines[0]).toContain("  a.txt");
    expect(lines[1]).toContain("  b.txt");
    const ha = createHash("sha256").update("a").digest("hex");
    expect(lines[0]!.startsWith(ha)).toBe(true);
    const info = buildBuildInfo({
      gitSha: "abc",
      packageVersion: "0.1.0",
      manifestVersion: "0.1.0.0",
      baseUrl: "https://example.com/excel-addin",
      viteBase: "/excel-addin/",
    });
    expect(info).toEqual({
      gitSha: "abc",
      packageVersion: "0.1.0",
      manifestVersion: "0.1.0.0",
      baseUrl: "https://example.com/excel-addin",
      viteBase: "/excel-addin/",
    });
    expect(JSON.stringify(info)).not.toMatch(/apiKey|secret|password/i);
  });

  it("makeArtifactName is deterministic and safe", () => {
    expect(makeArtifactName("0.1.0.0", "a5719f92f1a92e76527c8d75a895e79bb9dcba73")).toBe(
      "excel-addin-0.1.0.0-a5719f9",
    );
  });

  it("prod manifest for package base validates and uses same base", () => {
    const template = readFileSync(
      path.join(root, "manifest/templates/office-excel-manifest.template.xml"),
      "utf8",
    );
    const xml = renderOfficeManifest({
      mode: "prod",
      baseUrl: "https://example.com/excel-addin",
      version: "0.1.0.0",
      template,
    });
    const v = validateOfficeManifest(xml, { mode: "prod" });
    expect(v.ok, v.errors.join("; ")).toBe(true);
    expect(xml).toContain("https://example.com/excel-addin/index.html");
    expect(xml).toContain("https://example.com/excel-addin/assets/icon-32.png");
    expect(xml).not.toMatch(/localhost/);
  });
});
