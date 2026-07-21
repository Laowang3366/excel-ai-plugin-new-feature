import { createHash } from "node:crypto";
import {
  copyFileSync,
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
import { createWpsPackage } from "../scripts/package-wps-jsa.mjs";
import {
  makeWpsArtifactName,
  normalizeWpsGitSha,
  prepareWpsIndexHtml,
  renderWpsPublishXml,
  validateWpsEntryScript,
  validateWpsIndexHtml,
  validateWpsManifest,
  validateWpsPublishXml,
  validateWpsRibbon,
  validateWpsSourceBundle,
  LEGACY_OWN_ADDON_DIRECTORY,
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
  WPS_ENTRY_SCRIPT,
  WPS_PUBLISH_URL,
} from "../scripts/wpsJsaPackage.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceDir = path.join(root, "manifest/wps-jsa");
const tempRoots: string[] = [];

function makeTempRoot() {
  const dir = mkdtempSync(path.join(tmpdir(), "wengge-wps-jsa-"));
  tempRoots.push(dir);
  return dir;
}

function readSource(name: string) {
  return readFileSync(path.join(sourceDir, name), "utf8");
}

function seedWpsProject(project: string) {
  mkdirSync(path.join(project, "manifest/wps-jsa"), { recursive: true });
  for (const name of ["manifest.xml", "ribbon.xml", "publish.xml", WPS_ENTRY_SCRIPT]) {
    copyFileSync(path.join(sourceDir, name), path.join(project, "manifest/wps-jsa", name));
  }
  writeFileSync(path.join(project, "package.json"), '{"version":"0.1.0"}\n');
}

function seedBuiltDist(dist: string, indexHtml?: string) {
  mkdirSync(path.join(dist, "assets"), { recursive: true });
  writeFileSync(path.join(dist, "assets/app.js"), "console.log('app');\n");
  writeFileSync(path.join(dist, "assets/app.css"), "body{}\n");
  for (const size of [16, 32, 64, 80]) {
    copyFileSync(
      path.join(root, `public/assets/icon-${size}.png`),
      path.join(dist, `assets/icon-${size}.png`),
    );
  }
  writeFileSync(
    path.join(dist, "index.html"),
    indexHtml ??
      `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script type="module" crossorigin src="./assets/app.js"></script>
    <link rel="stylesheet" crossorigin href="./assets/app.css">
  </head>
  <body><div id="root"></div></body>
</html>
`,
  );
}

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("WPS JSA source validation", () => {
  it("accepts checked-in manifest/ribbon/entry/publish bundle", () => {
    const result = validateWpsSourceBundle({
      manifestXml: readSource("manifest.xml"),
      ribbonXml: readSource("ribbon.xml"),
      entryScript: readSource(WPS_ENTRY_SCRIPT),
      publishXml: readSource("publish.xml"),
    });
    expect(result).toEqual({ ok: true, errors: [] });
  });
  it("rejects direct image attribute; requires shared getImage callback", () => {
    const ribbon = readSource("ribbon.xml").replaceAll(
      'getImage="WenggeExcelAiGetImage"',
      'image="assets/icon-32.png"',
    );
    const res = validateWpsRibbon(ribbon);
    expect(res.ok).toBe(false);
    expect(res.errors.join(" ")).toMatch(/getImage|image/i);
    expect(validateWpsEntryScript(readSource(WPS_ENTRY_SCRIPT)).ok).toBe(true);
  });

  it("rejects tab getVisible (host-proven: omit so tab stays default-visible)", () => {
    const ribbon = readSource("ribbon.xml").replace(
      '<tab id="wenggeExcelAiTab" label="文格 AI">',
      '<tab id="wenggeExcelAiTab" label="文格 AI" getVisible="WenggeExcelAiTabVisible">',
    );
    const res = validateWpsRibbon(ribbon);
    expect(res.ok).toBe(false);
    expect(res.errors.join(" ")).toMatch(/getVisible/i);
    // checked-in ribbon must not declare getVisible
    expect(readSource("ribbon.xml")).not.toMatch(/getVisible/i);
  });


  it("matches desktop-style publish contract (jsplugins/et/file jsaddons url)", () => {
    const publish = renderWpsPublishXml();
    expect(publish).toContain("<jsplugins>");
    expect(publish).toContain(`name="${WPS_ADDON_NAME}"`);
    expect(publish).toContain('type="et"');
    expect(publish).toContain('enable="enable_dev"');
    expect(publish).toContain(WPS_PUBLISH_URL);
    expect(WPS_PUBLISH_URL).toBe(
      `file://%AppData%/kingsoft/wps/jsaddons/${WPS_ADDON_DIRECTORY}/index.html`,
    );
    expect(validateWpsPublishXml(publish)).toEqual({ ok: true, errors: [] });
  });

  it("rejects unsafe or drifted publish/manifest/ribbon/entry content", () => {
    expect(validateWpsManifest("<JsPlugin></JsPlugin>").ok).toBe(false);
    expect(
      validateWpsRibbon(
        `<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon/></customUI>`,
      ).ok,
    ).toBe(false);
    expect(validateWpsEntryScript("window.other = function () {}").ok).toBe(false);
    expect(validateWpsEntryScript("window.WenggeExcelAiOnLoad = function () {}; eval('1')").ok).toBe(
      false,
    );
    expect(
      validateWpsPublishXml(
        renderWpsPublishXml().replace(WPS_PUBLISH_URL, "https://evil.example/x"),
      ).ok,
    ).toBe(false);
    expect(
      validateWpsPublishXml(
        renderWpsPublishXml().replace(WPS_PUBLISH_URL, WPS_PUBLISH_URL.replace(WPS_ADDON_DIRECTORY, "../x")),
      ).ok,
    ).toBe(false);
    expect(
      validateWpsSourceBundle({
        manifestXml: readSource("manifest.xml"),
        ribbonXml: readSource("ribbon.xml"),
        entryScript: readSource(WPS_ENTRY_SCRIPT),
        publishXml: renderWpsPublishXml().replace(WPS_ADDON_NAME, "Other"),
      }).errors,
    ).toEqual(expect.arrayContaining([expect.stringMatching(/mismatch|drift/i)]));
  });
});

describe("WPS index.html Office.js removal and entry injection", () => {
  const withOffice = `<!doctype html><html><head>
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    <script type="module" src="./assets/app.js"></script>
  </head><body></body></html>`;

  it("removes Office.js CDN and injects relative wps-entry.js once", () => {
    const prepared = prepareWpsIndexHtml(withOffice);
    expect(prepared).not.toContain("appsforoffice.microsoft.com");
    expect(prepared).toContain(`src="./${WPS_ENTRY_SCRIPT}"`);
    const validation = validateWpsIndexHtml(prepared);
    expect(validation.ok).toBe(true);
    expect(validation.assets).toEqual(
      expect.arrayContaining(["assets/app.js", WPS_ENTRY_SCRIPT]),
    );
  });

  it("rejects missing/duplicate Office.js, prior injection, and root-absolute assets", () => {
    expect(() => prepareWpsIndexHtml("<html><head></head></html>")).toThrow(
      /exactly one Office\.js/,
    );
    expect(() =>
      prepareWpsIndexHtml(
        withOffice.replace(
          "office.js",
          "office.js\"></script><script src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js",
        ),
      ),
    ).toThrow(/exactly one Office\.js/);
    expect(() =>
      prepareWpsIndexHtml(
        withOffice.replace("</head>", `<script src="./${WPS_ENTRY_SCRIPT}"></script></head>`),
      ),
    ).toThrow(/already injected/);
    const absolute = prepareWpsIndexHtml(
      withOffice.replace("./assets/app.js", "/assets/app.js"),
    );
    const validation = validateWpsIndexHtml(absolute);
    expect(validation.ok).toBe(false);
    expect(validation.errors.join(" ")).toMatch(/root-absolute/);
  });

  it("rejects path traversal and non-allowlisted external assets in final index", () => {
    const prepared = prepareWpsIndexHtml(
      withOffice.replace("./assets/app.js", "./assets/../escape.js"),
    );
    const validation = validateWpsIndexHtml(prepared);
    expect(validation.ok).toBe(false);
    expect(validation.errors.join(" ")).toMatch(/traversal|not under|escape/i);
  });
});

describe("WPS package naming and git sha hardening", () => {
  it("normalizes git sha and builds artifact names", () => {
    expect(normalizeWpsGitSha("0123456789abcdef")).toBe("0123456789abcdef");
    expect(normalizeWpsGitSha("deadbeef")).toBe("deadbeef");
    expect(normalizeWpsGitSha("../evil")).toBe("unknown");
    expect(normalizeWpsGitSha("")).toBe("unknown");
    expect(makeWpsArtifactName("0.1.0", "0123456789abcdef")).toBe(
      "excel-addin-wps-jsa-0.1.0-0123456",
    );
    expect(() => makeWpsArtifactName("v0.1", "abc")).toThrow(/version/);
  });
});

describe("createWpsPackage layout, hashes, and path safety", () => {
  it("assembles local jsaddons layout without Office.js and with publish at package root", () => {
    const project = makeTempRoot();
    const dist = path.join(project, "dist");
    seedWpsProject(project);
    seedBuiltDist(dist);

    const summary = createWpsPackage({
      rootDir: project,
      distDir: dist,
      skipBuild: true,
      gitSha: "0123456789abcdef",
    });

    expect(summary.ok).toBe(true);
    expect(summary.artifactName).toBe("excel-addin-wps-jsa-0.1.0-0123456");
    expect(summary.addonDirectory).toBe(WPS_ADDON_DIRECTORY);
    expect(WPS_ADDON_DIRECTORY).toBe("WenggeExcelAiAddin_");
    expect(summary.addonDirectory).not.toBe(LEGACY_OWN_ADDON_DIRECTORY);
    expect(existsSync(path.join(dist, LEGACY_OWN_ADDON_DIRECTORY))).toBe(false);
    expect(existsSync(path.join(dist, WPS_ADDON_DIRECTORY, "index.html"))).toBe(true);

    expect(summary.files).toEqual(
      expect.arrayContaining([
        "publish.xml",
        "BUILD_INFO.json",
        "SHA256SUMS.txt",
        `${WPS_ADDON_DIRECTORY}/index.html`,
        `${WPS_ADDON_DIRECTORY}/manifest.xml`,
        `${WPS_ADDON_DIRECTORY}/ribbon.xml`,
        `${WPS_ADDON_DIRECTORY}/${WPS_ENTRY_SCRIPT}`,
        `${WPS_ADDON_DIRECTORY}/assets/app.js`,
        `${WPS_ADDON_DIRECTORY}/assets/icon-16.png`,
      ]),
    );

    const index = readFileSync(path.join(dist, WPS_ADDON_DIRECTORY, "index.html"), "utf8");
    expect(index).not.toContain("appsforoffice.microsoft.com");
    expect(index).toContain(`src="./${WPS_ENTRY_SCRIPT}"`);
    expect(index).toContain('src="./assets/app.js"');
    const entryAt = index.indexOf(`src="./${WPS_ENTRY_SCRIPT}"`);
    const moduleAt = index.indexOf('type="module"');
    expect(entryAt).toBeGreaterThanOrEqual(0);
    expect(moduleAt).toBeGreaterThan(entryAt);

    const publish = readFileSync(path.join(dist, "publish.xml"), "utf8");
    expect(publish).toContain(WPS_PUBLISH_URL);
    expect(existsSync(path.join(dist, WPS_ADDON_DIRECTORY, "publish.xml"))).toBe(false);

    const info = JSON.parse(readFileSync(path.join(dist, "BUILD_INFO.json"), "utf8")) as {
      target: string;
      gitSha: string;
      assetBase: string;
      addonDirectory: string;
    };
    expect(info).toMatchObject({
      target: "wps-jsa",
      gitSha: "0123456789abcdef",
      assetBase: "./",
      addonDirectory: WPS_ADDON_DIRECTORY,
    });
    expect(JSON.stringify(info)).not.toMatch(/api[_-]?key|secret|token/i);

    const sums = readFileSync(path.join(dist, "SHA256SUMS.txt"), "utf8").trim().split("\n");
    expect(sums.some((line) => line.endsWith(`  ${WPS_ADDON_DIRECTORY}/index.html`))).toBe(true);
    const indexHash = createHash("sha256").update(index).digest("hex");
    expect(sums.find((line) => line.endsWith(`  ${WPS_ADDON_DIRECTORY}/index.html`))).toBe(
      `${indexHash}  ${WPS_ADDON_DIRECTORY}/index.html`,
    );
    expect(sums.some((line) => line.includes("SHA256SUMS.txt"))).toBe(false);
  });

  it("rejects sensitive paths and removes dist after failed packaging", () => {
    const project = makeTempRoot();
    const dist = path.join(project, "dist");
    seedWpsProject(project);
    seedBuiltDist(dist);
    writeFileSync(path.join(dist, ".env"), "SECRET=1\n");

    expect(() =>
      createWpsPackage({
        rootDir: project,
        distDir: dist,
        skipBuild: true,
        gitSha: "0123456789abcdef",
      }),
    ).toThrow(/sensitive/);
    expect(existsSync(dist)).toBe(false);
  });

  it("rejects dist outside project or through symlink parents", () => {
    const parent = makeTempRoot();
    const project = path.join(parent, "project");
    const outside = path.join(parent, "outside");
    const outsideDist = path.join(outside, "dist");
    seedWpsProject(project);
    mkdirSync(outsideDist, { recursive: true });
    writeFileSync(path.join(outsideDist, "sentinel.txt"), "keep");
    seedBuiltDist(outsideDist);
    // win32: junction (absolute target) works without Developer Mode; other platforms use dir symlinks.
    const linked = path.join(project, "linked");
    const outsideAbs = path.resolve(outside);
    symlinkSync(
      outsideAbs,
      linked,
      process.platform === "win32" ? "junction" : "dir",
    );

    expect(() =>
      createWpsPackage({
        rootDir: project,
        distDir: path.join(project, "linked", "dist"),
        skipBuild: true,
      }),
    ).toThrow(/symlink|inside project root/);
    expect(existsSync(path.join(outsideDist, "sentinel.txt"))).toBe(true);

    expect(() =>
      createWpsPackage({
        rootDir: project,
        distDir: outsideDist,
        skipBuild: true,
      }),
    ).toThrow(/inside project root/);
  });

  it("rejects missing referenced assets and source bundle drift during package", () => {
    const project = makeTempRoot();
    const dist = path.join(project, "dist");
    seedWpsProject(project);
    seedBuiltDist(
      dist,
      `<!doctype html><html><head>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
        <script src="./assets/missing.js"></script>
      </head><body></body></html>`,
    );
    expect(() =>
      createWpsPackage({
        rootDir: project,
        distDir: dist,
        skipBuild: true,
      }),
    ).toThrow(/missing/);
    expect(existsSync(dist)).toBe(false);

    seedBuiltDist(dist);
    writeFileSync(
      path.join(project, "manifest/wps-jsa/publish.xml"),
      renderWpsPublishXml().replace(WPS_ADDON_NAME, "Tampered"),
    );
    expect(() =>
      createWpsPackage({
        rootDir: project,
        distDir: dist,
        skipBuild: true,
      }),
    ).toThrow(/mismatch|drift/i);
  });
});
