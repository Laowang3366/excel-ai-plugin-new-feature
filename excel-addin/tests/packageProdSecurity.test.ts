import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  symlinkSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, describe, expect, it } from "vitest";
import {
  assertIndexAssetsUnderBase,
  assertLocalAssetFiles,
  formatSpawnFailure,
  listFilesRecursiveStrict,
} from "../scripts/packageProdCore.mjs";
import { createPackage } from "../scripts/package-prod.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const template = readFileSync(
  path.join(root, "manifest/templates/office-excel-manifest.template.xml"),
  "utf8",
);
const tempRoots: string[] = [];

function makeTempRoot() {
  const dir = mkdtempSync(path.join(tmpdir(), "wengge-package-prod-"));
  tempRoots.push(dir);
  return dir;
}

function indexWith(ref: string, tag: "script" | "link" = "script") {
  return tag === "script"
    ? `<!doctype html><script src="${ref}"></script>`
    : `<!doctype html><link rel="stylesheet" href="${ref}">`;
}

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("production package asset URL hardening", () => {
  it("normalizes valid absolute/relative assets and allows only Office.js CDN", () => {
    const html = `<!doctype html>
      <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
      <script src="/excel-addin/assets/app.js"></script>
      <link rel="stylesheet" href="assets/app.css?build=1#x">`;
    expect(assertIndexAssetsUnderBase(html, "/excel-addin/")).toEqual([
      "assets/app.js",
      "assets/app.css",
    ]);
  });

  it.each([
    ["../escape.js", /traversal|not under VITE_BASE/],
    ["/excel-addin/../other/x.js", /traversal/],
    ["/excel-addin/%2e%2e/other/x.js", /traversal/],
    ["/excel-addin/%2Fother/x.js", /encoded path separators/],
    ["/app/%252e%252e/escape.js", /nested percent encoding/],
    ["/app/%252fescape.js", /nested percent encoding/],
    ["/app/%255cescape.js", /nested percent encoding/],
    ["/app/%25%32%65%25%32%65/escape.js", /nested percent encoding/],
    ["//evil.example/x.js", /protocol-relative/],
    ["https://evil.example/x.js", /not allowlisted/],
    ["https://appsforoffice.microsoft.com/lib/1/hosted/office.js?x=1", /not allowlisted/],
    ["data:text/javascript,alert(1)", /unsupported asset URL protocol/],
    ["javascript:alert(1)", /unsupported asset URL protocol/],
    ["file:///tmp/x.js", /unsupported asset URL protocol/],
    ["blob:https://example.com/id", /unsupported asset URL protocol/],
    ["assets\\evil.js", /backslash/],
    ["assets/%00evil.js", /control/],
    ["/app2/assets/x.js", /not under VITE_BASE/],
  ])("rejects unsafe asset ref %s", (ref, expected) => {
    expect(() => assertIndexAssetsUnderBase(indexWith(ref), "/app/")).toThrow(
      expected,
    );
  });

  it("rejects unquoted asset attributes", () => {
    expect(() =>
      assertIndexAssetsUnderBase("<script src=/app/assets/x.js></script>", "/app/"),
    ).toThrow(/quoted URL/);
    expect(() =>
      assertIndexAssetsUnderBase("<script src></script>", "/app/"),
    ).toThrow(/quoted URL/);
  });

  it("matches only exact attributes and rejects duplicate targets", () => {
    expect(
      assertIndexAssetsUnderBase(
        '<script data-src="/app/assets/ignored.js"></script>',
        "/app/",
      ),
    ).toEqual([]);
    expect(() =>
      assertIndexAssetsUnderBase(
        '<script src="/app/assets/a.js" src="/app/assets/b.js"></script>',
        "/app/",
      ),
    ).toThrow(/duplicate/);
  });

  it.each([
    [
      '<script data-src="/app/assets/good.js" src=//evil.example/x.js></script>',
      /quoted URL/,
    ],
    [
      '<script src=//evil.example/x.js src="/app/assets/good.js"></script>',
      /quoted URL|duplicate/,
    ],
    [
      '<link data-href="/app/assets/good.css" href=//evil.example/x.css>',
      /quoted URL/,
    ],
  ])("rejects target-attribute confusion in %s", (html, expected) => {
    expect(() => assertIndexAssetsUnderBase(html, "/app/")).toThrow(expected);
  });

  it("rejects HTML character references before URL resolution", () => {
    expect(() =>
      assertIndexAssetsUnderBase(
        '<script src="/app/&period;&period;/escape.js"></script>',
        "/app/",
      ),
    ).toThrow(/character references/);
  });

  it("does not hide an unsafe URL behind a greater-than sign in another attribute", () => {
    expect(() =>
      assertIndexAssetsUnderBase(
        '<script data-note=">" src="//evil.example/x.js"></script>',
        "/app/",
      ),
    ).toThrow(/protocol-relative/);
  });
});

describe("production package filesystem hardening", () => {
  it("requires referenced assets to be regular files inside dist", () => {
    const dist = makeTempRoot();
    mkdirSync(path.join(dist, "assets"));
    writeFileSync(path.join(dist, "assets/app.js"), "ok");
    mkdirSync(path.join(dist, "assets/directory"));

    expect(assertLocalAssetFiles(dist, ["assets/app.js"])).toHaveLength(1);
    expect(() => assertLocalAssetFiles(dist, ["assets/missing.js"])).toThrow(
      /missing/,
    );
    expect(() => assertLocalAssetFiles(dist, ["assets/directory"])).toThrow(
      /not a regular file/,
    );
    expect(() => assertLocalAssetFiles(dist, ["../outside.js"])).toThrow(
      /invalid local asset path/,
    );
  });

  it.skipIf(process.platform === "win32")(
    "rejects file and directory symlinks without following them",
    () => {
      const parent = makeTempRoot();
      const dist = path.join(parent, "dist");
      const outside = path.join(parent, "outside");
      mkdirSync(dist);
      mkdirSync(outside);
      writeFileSync(path.join(outside, "external.js"), "outside");
      symlinkSync(path.join(outside, "external.js"), path.join(dist, "linked.js"));
      symlinkSync(outside, path.join(dist, "linked-dir"), "dir");

      expect(() => assertLocalAssetFiles(dist, ["linked.js"])).toThrow(/symlink/);
      expect(() => listFilesRecursiveStrict(dist)).toThrow(/symlink/);
    },
  );

  it("removes dist when post-build package validation fails", () => {
    const project = makeTempRoot();
    const dist = path.join(project, "dist");
    mkdirSync(path.join(project, "manifest/templates"), { recursive: true });
    mkdirSync(dist);
    writeFileSync(path.join(project, "package.json"), '{"version":"0.1.0"}\n');
    writeFileSync(
      path.join(project, "manifest/templates/office-excel-manifest.template.xml"),
      template,
    );
    writeFileSync(
      path.join(dist, "index.html"),
      indexWith("/excel-addin/assets/missing.js"),
    );

    expect(() =>
      createPackage({
        baseUrl: "https://example.com/excel-addin",
        rootDir: project,
        distDir: dist,
        skipBuild: true,
      }),
    ).toThrow(/missing/);
    expect(existsSync(dist)).toBe(false);
  });

  it("does not delete outside the project through a linked dist parent", () => {
    const parent = makeTempRoot();
    const project = path.join(parent, "project");
    const outside = path.join(parent, "outside");
    const outsideDist = path.join(outside, "dist");
    mkdirSync(path.join(project, "manifest/templates"), { recursive: true });
    mkdirSync(outsideDist, { recursive: true });
    writeFileSync(path.join(outsideDist, "sentinel.txt"), "keep");
    writeFileSync(path.join(project, "package.json"), '{"version":"0.1.0"}\n');
    writeFileSync(
      path.join(project, "manifest/templates/office-excel-manifest.template.xml"),
      template,
    );
    symlinkSync(outside, path.join(project, "linked"), "junction");

    expect(() =>
      createPackage({
        baseUrl: "https://example.com/app",
        rootDir: project,
        distDir: path.join(project, "linked/dist"),
        skipBuild: true,
      }),
    ).toThrow(/symlink/);
    expect(existsSync(path.join(outsideDist, "sentinel.txt"))).toBe(true);
  });

  it("formats spawn start, signal, and exit failures safely", () => {
    expect(formatSpawnFailure({ error: { code: "ENOENT" } })).toBe(
      "npm run build failed to start ENOENT",
    );
    expect(formatSpawnFailure({ signal: "SIGTERM", status: null })).toBe(
      "npm run build terminated by signal SIGTERM",
    );
    expect(formatSpawnFailure({ status: 2 })).toBe(
      "npm run build failed with status 2",
    );
  });
});
