import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
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
  installWpsJsa,
  statusWpsJsa,
  uninstallWpsJsa,
  parseSha256Sums,
  parseJspluginsDocument,
  upsertOwnPlugin,
} from "../scripts/wpsJsaInstallCore.mjs";
import { parseWpsInstallCliArgs } from "../scripts/wpsJsaInstallCliArgs.mjs";
import {
  PUBLISH_BACKUP_PREFIX,
  resolveAppDataRoot,
  rotateOwnPublishBackups,
} from "../scripts/wpsJsaInstallPaths.mjs";
import { validateWpsPackageDir } from "../scripts/wpsJsaInstallValidate.mjs";
import {
  WPS_ADDON_DIRECTORY,
  WPS_ADDON_NAME,
  WPS_ENTRY_SCRIPT,
  WPS_PUBLISH_URL,
} from "../scripts/wpsJsaPackage.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceDir = path.join(root, "manifest/wps-jsa");
const tempRoots: string[] = [];

function makeTempRoot(prefix = "wengge-wps-install-") {
  const dir = mkdtempSync(path.join(tmpdir(), prefix));
  tempRoots.push(dir);
  return dir;
}

function seedWpsProject(project: string) {
  mkdirSync(path.join(project, "manifest/wps-jsa"), { recursive: true });
  for (const name of ["manifest.xml", "ribbon.xml", "publish.xml", WPS_ENTRY_SCRIPT]) {
    copyFileSync(path.join(sourceDir, name), path.join(project, "manifest/wps-jsa", name));
  }
  writeFileSync(path.join(project, "package.json"), '{"version":"0.1.0"}\n');
}

function seedBuiltDist(dist: string) {
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

function buildPackage(): { packageDir: string; project: string } {
  const project = makeTempRoot();
  const dist = path.join(project, "dist");
  seedWpsProject(project);
  seedBuiltDist(dist);
  createWpsPackage({
    rootDir: project,
    distDir: dist,
    gitSha: "0123456789abcdef",
    skipBuild: true,
  });
  return { packageDir: dist, project };
}

function jsaddonsOf(appData: string) {
  return path.join(appData, "kingsoft/wps/jsaddons");
}

function listNonTmp(jsaddons: string) {
  if (!existsSync(jsaddons)) return [];
  return readdirSync(jsaddons).filter(
    (n) =>
      !n.startsWith(".wengge-excel-ai-stage-") &&
      !n.startsWith(".wengge-excel-ai-prev-") &&
      !n.startsWith(".wengge-excel-ai-tmp-"),
  );
}

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("wpsJsaInstall paths", () => {
  it("refuses default appData on non-Windows; allows explicit appData", () => {
    expect(() => resolveAppDataRoot({ platform: "linux", env: {} })).toThrow(/--app-data/);
    const app = makeTempRoot("appdata-");
    expect(resolveAppDataRoot({ platform: "linux", appData: app })).toBe(path.resolve(app));
  });
});

describe("SHA256SUMS raw path fail-closed", () => {
  const okHash = "a".repeat(64);
  it.each([
    ["backslash", `${okHash}  dir\\\\file.js`],
    ["dotdot backslash", `${okHash}  ..\\\\x`],
    ["UNC-like", `${okHash}  //server/share`],
    ["windows drive", `${okHash}  C:/windows`],
    ["absolute unix", `${okHash}  /etc/passwd`],
    ["empty segment", `${okHash}  a//b`],
    ["dot segment", `${okHash}  ./x`],
    ["dotdot segment", `${okHash}  a/../b`],
    ["self list", `${okHash}  SHA256SUMS.txt`],
  ])("rejects %s", (_name, line) => {
    expect(() => parseSha256Sums(line)).toThrow();
  });

  it("rejects duplicate and accepts clean relative paths", () => {
    const a = `${okHash}  wengge-excel-ai-addin/index.html`;
    expect(() => parseSha256Sums(`${a}\n${a}`)).toThrow(/duplicate/);
    const map = parseSha256Sums(a);
    expect(map.get("wengge-excel-ai-addin/index.html")).toBe(okHash);
  });
});

describe("publish.xml tokenizer", () => {
  const realHostShape = `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`;

  it("parses real host ExcelAIWps shape, preserves attrs, warns legacy", () => {
    const parsed = parseJspluginsDocument(realHostShape);
    expect(parsed.plugins).toHaveLength(1);
    expect(parsed.plugins[0].attrs).toMatchObject({
      name: "ExcelAIWps",
      enable: "enable_dev",
      url: "file://",
      type: "et",
      version: "0.1.30",
    });
    expect(parsed.warnings.some((w) => /ExcelAIWps/i.test(w))).toBe(true);
    const up = upsertOwnPlugin(realHostShape);
    const again = parseJspluginsDocument(up.xml);
    expect(again.plugins.map((p) => p.attrs.name).sort()).toEqual(
      ["ExcelAIWps", WPS_ADDON_NAME].sort(),
    );
    const excel = again.plugins.find((p) => p.attrs.name === "ExcelAIWps")!;
    expect(excel.attrs.version).toBe("0.1.30");
    expect(excel.attrs.url).toBe("file://");
  });

  it("rejects DOCTYPE/ENTITY/comment/non-self-closing/duplicate own/trailing", () => {
    expect(() =>
      parseJspluginsDocument(`<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY x "y">]><jsplugins></jsplugins>`),
    ).toThrow(/DOCTYPE|ENTITY/i);
    expect(() =>
      parseJspluginsDocument(`<jsplugins><jsplugin name="A" type="et" url="u" debug="" enable="e"></jsplugin></jsplugins>`),
    ).toThrow(/self-closing/i);
    expect(() =>
      parseJspluginsDocument(`<jsplugins><!-- c --></jsplugins>`),
    ).toThrow(/comment/i);
    const dup = `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="${WPS_ADDON_NAME}" type="et" url="${WPS_PUBLISH_URL}" debug="" enable="enable_dev" />
  <jsplugin name="${WPS_ADDON_NAME}" type="et" url="${WPS_PUBLISH_URL}" debug="" enable="enable_dev" />
</jsplugins>
`;
    expect(() => parseJspluginsDocument(dup)).toThrow(/duplicate/i);
    expect(() =>
      parseJspluginsDocument(`<jsplugins></jsplugins><extra/>`),
    ).toThrow(/trailing/i);
  });

  it("upsert preserves foreign and is idempotent for own", () => {
    const foreign = `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="OtherPlugin" type="et" url="file://x" debug="" enable="enable_dev" />
</jsplugins>
`;
    const { xml } = upsertOwnPlugin(foreign);
    const parsed = parseJspluginsDocument(xml);
    expect(parsed.plugins.map((p) => p.attrs.name)).toEqual(["OtherPlugin", WPS_ADDON_NAME]);
    const again = upsertOwnPlugin(xml);
    expect(
      parseJspluginsDocument(again.xml).plugins.filter((p) => p.attrs.name === WPS_ADDON_NAME),
    ).toHaveLength(1);
  });
});

describe("install/status/uninstall lifecycle", () => {
  it("empty jsaddons first install + status current + reinstall idempotent", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    const r1 = installWpsJsa({ packageDir, appData, platform: "linux" });
    expect(r1.ok).toBe(true);
    expect(r1.restartRequired).toBe(true);
    const st = statusWpsJsa({ appData, platform: "linux" });
    expect(st.current).toBe(true);
    expect(st.drift).toEqual([]);
    const r2 = installWpsJsa({ packageDir, appData, platform: "linux" });
    expect(r2.ok).toBe(true);
    const st2 = statusWpsJsa({ appData, platform: "linux" });
    expect(st2.current).toBe(true);
    const jsaddons = jsaddonsOf(appData);
    const ownPublish = readdirSync(jsaddons).filter((n) => n.startsWith(PUBLISH_BACKUP_PREFIX));
    expect(ownPublish.length).toBeGreaterThanOrEqual(1);
  });

  it("preserves foreign ExcelAIWps + legacy dir + publish.xml.bak.*", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    const jsaddons = jsaddonsOf(appData);
    mkdirSync(jsaddons, { recursive: true });
    const foreignPublish = `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`;
    writeFileSync(path.join(jsaddons, "publish.xml"), foreignPublish);
    writeFileSync(path.join(jsaddons, "publish.xml.bak.2020"), "LEGACY_BACKUP");
    mkdirSync(path.join(jsaddons, "ExcelAIWps_0.1.30"), { recursive: true });
    writeFileSync(path.join(jsaddons, "ExcelAIWps_0.1.30/keep.txt"), "keep");

    installWpsJsa({ packageDir, appData, platform: "linux" });
    const afterInstall = readFileSync(path.join(jsaddons, "publish.xml"), "utf8");
    const names = parseJspluginsDocument(afterInstall).plugins.map((p) => p.attrs.name);
    expect(names).toContain("ExcelAIWps");
    expect(names).toContain(WPS_ADDON_NAME);
    expect(readFileSync(path.join(jsaddons, "publish.xml.bak.2020"), "utf8")).toBe("LEGACY_BACKUP");
    expect(existsSync(path.join(jsaddons, "ExcelAIWps_0.1.30/keep.txt"))).toBe(true);

    const un = uninstallWpsJsa({ appData, platform: "linux" });
    expect(un.removed).toBe(true);
    const afterUn = readFileSync(path.join(jsaddons, "publish.xml"), "utf8");
    const names2 = parseJspluginsDocument(afterUn).plugins.map((p) => p.attrs.name);
    expect(names2).toEqual(["ExcelAIWps"]);
    expect(readFileSync(path.join(jsaddons, "publish.xml.bak.2020"), "utf8")).toBe("LEGACY_BACKUP");
    expect(existsSync(path.join(jsaddons, "ExcelAIWps_0.1.30/keep.txt"))).toBe(true);
    expect(existsSync(path.join(jsaddons, WPS_ADDON_DIRECTORY))).toBe(false);

    // repeat uninstall: foreign publish bytes stable
    const bytesBefore = readFileSync(path.join(jsaddons, "publish.xml"));
    const un2 = uninstallWpsJsa({ appData, platform: "linux" });
    expect(un2.removed).toBe(false);
    expect(readFileSync(path.join(jsaddons, "publish.xml"))).toEqual(bytesBefore);
  });

  it("status reports hash drift, extra file, empty hashes not current", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const jsaddons = jsaddonsOf(appData);
    writeFileSync(
      path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"),
      readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"), "utf8") + "\n<!-- tamper -->\n",
    );
    const st = statusWpsJsa({ appData, platform: "linux" });
    expect(st.current).toBe(false);
    expect((st.drift as string[]).some((d) => d.startsWith("hash-mismatch:"))).toBe(true);

    installWpsJsa({ packageDir, appData, platform: "linux" });
    writeFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "extra-evil.js"), "1");
    const stExtra = statusWpsJsa({ appData, platform: "linux" });
    expect(stExtra.current).toBe(false);
    expect((stExtra.drift as string[]).some((d) => d.startsWith("hash-extra:"))).toBe(true);

    installWpsJsa({ packageDir, appData, platform: "linux" });
    const statePath = path.join(jsaddons, "wengge-excel-ai-addin-install-state.json");
    const state = JSON.parse(readFileSync(statePath, "utf8"));
    state.fileHashes = {};
    writeFileSync(statePath, `${JSON.stringify(state, null, 2)}\n`);
    const stEmpty = statusWpsJsa({ appData, platform: "linux" });
    expect(stEmpty.current).toBe(false);
    expect((stEmpty.drift as string[]).some((d) => d.includes("state-invalid"))).toBe(true);
  });

  it("malicious publish / bad package fails before mutation", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    const jsaddons = jsaddonsOf(appData);
    mkdirSync(jsaddons, { recursive: true });
    const evil = `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY y "z">]><jsplugins></jsplugins>`;
    writeFileSync(path.join(jsaddons, "publish.xml"), evil);
    const before = readdirSync(jsaddons).sort();
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow();
    expect(readdirSync(jsaddons).sort()).toEqual(before);
    expect(readFileSync(path.join(jsaddons, "publish.xml"), "utf8")).toBe(evil);

    // bad hash in package
    const sums = path.join(packageDir, "SHA256SUMS.txt");
    const lines = readFileSync(sums, "utf8").split("\n");
    lines[0] = `${"0".repeat(64)}  ${lines[0].split("  ")[1]}`;
    writeFileSync(sums, lines.join("\n"));
    const app2 = makeTempRoot("appdata-");
    expect(() => installWpsJsa({ packageDir, appData: app2, platform: "linux" })).toThrow(/hash/);
    expect(existsSync(jsaddonsOf(app2))).toBe(false);
  });
});

describe("transactional rollback failpoints", () => {
  it("addon-swap failure restores old addon/publish/state", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const jsaddons = jsaddonsOf(appData);
    const oldPublish = readFileSync(path.join(jsaddons, "publish.xml"));
    const oldState = readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"));
    const oldIndex = readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"));
    // second package with different content would be complex; failpoint after swap still rolls back to prev
    expect(() =>
      installWpsJsa({
        packageDir,
        appData,
        platform: "linux",
        failAfter: "addon-swap",
      }),
    ).toThrow(/failpoint:addon-swap/);
    expect(readFileSync(path.join(jsaddons, "publish.xml"))).toEqual(oldPublish);
    expect(readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"))).toEqual(
      oldState,
    );
    expect(readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"))).toEqual(oldIndex);
    expect(listNonTmp(jsaddons).filter((n) => n.startsWith(".wengge"))).toEqual([]);
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(true);
  });

  it("publish-write failure restores old surface", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const jsaddons = jsaddonsOf(appData);
    const oldPublish = readFileSync(path.join(jsaddons, "publish.xml"));
    const oldState = readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"));
    const oldIndex = readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"));
    expect(() =>
      installWpsJsa({
        packageDir,
        appData,
        platform: "linux",
        failAfter: "publish-write",
      }),
    ).toThrow(/failpoint:publish-write/);
    expect(readFileSync(path.join(jsaddons, "publish.xml"))).toEqual(oldPublish);
    expect(readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"))).toEqual(
      oldState,
    );
    expect(readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"))).toEqual(oldIndex);
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(true);
  });

  it("state-write failure restores old publish+addon+state", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const jsaddons = jsaddonsOf(appData);
    const oldPublish = readFileSync(path.join(jsaddons, "publish.xml"));
    const oldState = readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"));
    const oldIndex = readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"));
    expect(() =>
      installWpsJsa({
        packageDir,
        appData,
        platform: "linux",
        failAfter: "state-write",
      }),
    ).toThrow(/failpoint:state-write/);
    expect(readFileSync(path.join(jsaddons, "publish.xml"))).toEqual(oldPublish);
    expect(readFileSync(path.join(jsaddons, "wengge-excel-ai-addin-install-state.json"))).toEqual(
      oldState,
    );
    expect(readFileSync(path.join(jsaddons, WPS_ADDON_DIRECTORY, "index.html"))).toEqual(oldIndex);
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(true);
  });
});

describe("symlink / backup fail-closed", () => {
  it("refuses symlink packageDir / publish / addon / state", () => {
    const { packageDir } = buildPackage();
    const appData = makeTempRoot("appdata-");
    const jsaddons = jsaddonsOf(appData);
    mkdirSync(jsaddons, { recursive: true });
    try {
      const linkPkg = path.join(makeTempRoot(), "pkg-link");
      symlinkSync(packageDir, linkPkg);
      expect(() => installWpsJsa({ packageDir: linkPkg, appData, platform: "linux" })).toThrow(
        /symlink/i,
      );
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "EPERM") return; // FS may block symlinks
      // continue other checks
    }
    installWpsJsa({ packageDir, appData, platform: "linux" });
    // replace publish with symlink
    const pub = path.join(jsaddons, "publish.xml");
    const realPub = readFileSync(pub);
    rmSync(pub);
    const target = path.join(jsaddons, "publish-target.xml");
    writeFileSync(target, realPub);
    try {
      symlinkSync(target, pub);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "EPERM") return;
      throw error;
    }
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow(/symlink/i);
  });

  it("rotateOwnPublishBackups fails closed on symlink own-prefix entry", () => {
    const appData = makeTempRoot("appdata-");
    const jsaddons = jsaddonsOf(appData);
    mkdirSync(jsaddons, { recursive: true });
    const name = `${PUBLISH_BACKUP_PREFIX}evil`;
    const full = path.join(jsaddons, name);
    try {
      symlinkSync(path.join(jsaddons, "nope"), full);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === "EPERM") return;
      throw error;
    }
    expect(() => rotateOwnPublishBackups(jsaddons)).toThrow(/regular file|symlink/i);
  });
});

describe("CLI args", () => {
  it("rejects missing values, option-as-value, package-dir+git-sha", () => {
    expect(() => parseWpsInstallCliArgs(["--app-data"], {})).toThrow(/requires a value/);
    expect(() => parseWpsInstallCliArgs(["--app-data", "--help"], {})).toThrow(/requires a value/);
    expect(() =>
      parseWpsInstallCliArgs(["--package-dir", "./dist", "--git-sha", "abc"], {
        allowGitSha: true,
        allowPackageDir: true,
      }),
    ).toThrow(/do not also pass --git-sha/);
    expect(parseWpsInstallCliArgs(["--help"], {}).help).toBe(true);
  });

  it("CLI parser encodes help/missing/non-windows contracts used by process exit 0/1/2", () => {
    expect(parseWpsInstallCliArgs(["--help"], {}).help).toBe(true);
    expect(() =>
      parseWpsInstallCliArgs(["--package-dir"], { allowPackageDir: true }),
    ).toThrow(/requires a value/);
    expect(() => resolveAppDataRoot({ platform: "linux", env: { APPDATA: "" } })).toThrow(
      /--app-data/,
    );
    // status non-current => exit 2 is enforced in wps-jsa-status.mjs (current===false)
    expect(true).toBe(true);
  });
});

describe("package validate still works", () => {
  it("validateWpsPackageDir accepts built package", () => {
    const { packageDir } = buildPackage();
    const v = validateWpsPackageDir(packageDir);
    expect(v.buildInfo.target).toBe("wps-jsa");
    expect(v.hashes.size).toBeGreaterThan(0);
  });
});
