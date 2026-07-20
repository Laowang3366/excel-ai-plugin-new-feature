import { createHash } from "node:crypto";
import {
  copyFileSync,
  existsSync,
  mkdirSync,
  mkdtempSync,
  readFileSync,
  readdirSync,
  readlinkSync,
  renameSync,
  lstatSync,
  statSync,
  rmSync,
  symlinkSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
import { spawnSync } from "node:child_process";
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
  ownPluginMatchesContract,
  listActiveTempNames,
  PUBLISH_BACKUP_PREFIX,
  rotateOwnPublishBackups,
  resolveAppDataRoot,
} from "../scripts/wpsJsaInstallCore.mjs";
import { parseWpsInstallCliArgs } from "../scripts/wpsJsaInstallCliArgs.mjs";
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

function seedBuiltDist(dist: string, marker: string) {
  mkdirSync(path.join(dist, "assets"), { recursive: true });
  writeFileSync(path.join(dist, "assets/app.js"), `console.log(${JSON.stringify(marker)});\n`);
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
  <body><div id="root" data-marker="${marker}"></div></body>
</html>
`,
  );
}

function buildPackage(marker = "default"): { packageDir: string; marker: string } {
  const project = makeTempRoot();
  const dist = path.join(project, "dist");
  seedWpsProject(project);
  seedBuiltDist(dist, marker);
  createWpsPackage({
    rootDir: project,
    distDir: dist,
    gitSha: "0123456789abcdef",
    skipBuild: true,
  });
  return { packageDir: dist, marker };
}

function jsaddonsOf(appData: string) {
  return path.join(appData, "kingsoft/wps/jsaddons");
}

function snapshotTree(rootDir: string): {
  entries: string[];
  files: Record<string, { size: number; mtimeMs: number; sha: string }>;
} {
  const entries: string[] = [];
  const files: Record<string, { size: number; mtimeMs: number; sha: string }> = {};
  if (!existsSync(rootDir)) return { entries, files };
  const walk = (dir: string, rel = "") => {
    for (const name of readdirSync(dir).sort()) {
      const abs = path.join(dir, name);
      const r = rel ? `${rel}/${name}` : name;
      entries.push(r);
      const st = lstatSync(abs);
      if (st.isSymbolicLink()) {
        files[r] = {
          size: 0,
          mtimeMs: st.mtimeMs,
          sha: `symlink:${readlinkSync(abs)}`,
        };
        continue;
      }
      if (st.isDirectory()) walk(abs, r);
      else if (st.isFile()) {
        files[r] = {
          size: st.size,
          mtimeMs: st.mtimeMs,
          sha: createHash("sha256").update(readFileSync(abs)).digest("hex"),
        };
      }
    }
  };
  walk(rootDir);
  return { entries, files };
}

function addonMarker(appData: string) {
  const html = readFileSync(
    path.join(jsaddonsOf(appData), WPS_ADDON_DIRECTORY, "index.html"),
    "utf8",
  );
  const m = /data-marker="([^"]+)"/.exec(html);
  return m?.[1] ?? null;
}

function activeTemps(appData: string) {
  return listActiveTempNames(jsaddonsOf(appData));
}

function trySymlink(target: string, linkPath: string): "ok" | "skip" {
  try {
    symlinkSync(target, linkPath);
    return "ok";
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === "EPERM" || code === "EACCES" || code === "ENOTSUP") return "skip";
    throw error;
  }
}

afterEach(() => {
  for (const dir of tempRoots.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

describe("SHA256SUMS raw path fail-closed", () => {
  const okHash = "a".repeat(64);
  it.each([
    ["backslash", `${okHash}  dir\\\\file.js`],
    ["absolute", `${okHash}  /etc/passwd`],
    ["drive", `${okHash}  C:/x`],
    ["dotdot", `${okHash}  a/../b`],
  ])("rejects %s", (_n, line) => {
    expect(() => parseSha256Sums(line)).toThrow();
  });
});

describe("publish contract", () => {
  it("parses real ExcelAIWps host shape and preserves version", () => {
    const real = `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`;
    const parsed = parseJspluginsDocument(real);
    expect((parsed.plugins[0].attrs as Record<string, string>).version).toBe("0.1.30");
    const up = upsertOwnPlugin(real);
    const again = parseJspluginsDocument(up.xml);
    const excel = again.plugins.find(
      (p) => (p.attrs as Record<string, string>).name === "ExcelAIWps",
    );
    expect((excel?.attrs as Record<string, string> | undefined)?.version).toBe("0.1.30");
  });

  it("rejects XML declaration 1.1 / unknown attrs", () => {
    expect(() =>
      parseJspluginsDocument(`<?xml version="1.1"?><jsplugins></jsplugins>`),
    ).toThrow(/1\.0|declaration/i);
    expect(() =>
      parseJspluginsDocument(`<?xml version="1.0" standalone="yes"?><jsplugins></jsplugins>`),
    ).toThrow(/declaration/i);
  });

  it("ownPluginMatchesContract requires exact five attrs", () => {
    expect(
      ownPluginMatchesContract({
        name: WPS_ADDON_NAME,
        type: "et",
        url: WPS_PUBLISH_URL,
        debug: "",
        enable: "enable_dev",
      }),
    ).toBe(true);
    expect(
      ownPluginMatchesContract({
        name: WPS_ADDON_NAME,
        type: "et",
        url: WPS_PUBLISH_URL,
        enable: "enable_dev",
      }),
    ).toBe(false);
    expect(
      ownPluginMatchesContract({
        name: WPS_ADDON_NAME,
        type: "et",
        url: WPS_PUBLISH_URL,
        debug: "",
        enable: "enable_dev",
        extra: "x",
      }),
    ).toBe(false);
  });
});

describe("install lifecycle + TOCTOU", () => {
  it("installs and reports current", () => {
    const { packageDir } = buildPackage("m1");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const st = statusWpsJsa({ appData, platform: "linux" });
    expect(st.current).toBe(true);
    expect(addonMarker(appData)).toBe("m1");
    expect(activeTemps(appData)).toEqual([]);
  });

  it("afterValidate package extra fails before appData mutation", () => {
    const { packageDir } = buildPackage("base");
    const appData = makeTempRoot("app-");
    expect(() =>
      installWpsJsa({
        packageDir,
        appData,
        platform: "linux",
        afterValidate: (dir) => {
          writeFileSync(path.join(dir, WPS_ADDON_DIRECTORY, "extra-evil.js"), "evil\n");
        },
      }),
    ).toThrow();
    // jsaddons should not exist under appData (no mutation)
    expect(existsSync(jsaddonsOf(appData))).toBe(false);
  });
});

describe("install rollback with distinct old/new packages", () => {
  function snapshotInstall(appData: string) {
    const js = jsaddonsOf(appData);
    return {
      marker: addonMarker(appData),
      publish: readFileSync(path.join(js, "publish.xml")),
      state: readFileSync(path.join(js, "wengge-excel-ai-addin-install-state.json")),
      index: readFileSync(path.join(js, WPS_ADDON_DIRECTORY, "index.html")),
    };
  }

  it.each([
    "addon-swap",
    "publish-write",
    "publish-write-after",
    "state-write",
    "state-write-after",
  ] as const)("failAfter %s restores OLD marker not NEW", (fp) => {
    const oldPkg = buildPackage("OLD_MARKER");
    const newPkg = buildPackage("NEW_MARKER");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir: oldPkg.packageDir, appData, platform: "linux" });
    expect(addonMarker(appData)).toBe("OLD_MARKER");
    const before = snapshotInstall(appData);

    expect(() =>
      installWpsJsa({
        packageDir: newPkg.packageDir,
        appData,
        platform: "linux",
        failAfter: fp,
      }),
    ).toThrow(new RegExp(`failpoint:${fp}`));

    expect(addonMarker(appData)).toBe("OLD_MARKER");
    expect(readFileSync(path.join(jsaddonsOf(appData), "publish.xml"))).toEqual(before.publish);
    expect(
      readFileSync(path.join(jsaddonsOf(appData), "wengge-excel-ai-addin-install-state.json")),
    ).toEqual(before.state);
    expect(
      readFileSync(path.join(jsaddonsOf(appData), WPS_ADDON_DIRECTORY, "index.html")),
    ).toEqual(before.index);
    expect(activeTemps(appData)).toEqual([]);
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(true);
  });
});

describe("uninstall true transaction + state failure restores addon", () => {
  it("failAfter state restores addon/publish/state and status current", () => {
    const { packageDir } = buildPackage("KEEP");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const beforeMarker = addonMarker(appData);
    const js = jsaddonsOf(appData);
    const beforePublish = readFileSync(path.join(js, "publish.xml"));
    const beforeState = readFileSync(path.join(js, "wengge-excel-ai-addin-install-state.json"));
    const beforeIndex = readFileSync(path.join(js, WPS_ADDON_DIRECTORY, "index.html"));

    expect(() =>
      uninstallWpsJsa({ appData, platform: "linux", failAfter: "state" }),
    ).toThrow(/failpoint:state/);

    expect(existsSync(path.join(js, WPS_ADDON_DIRECTORY))).toBe(true);
    expect(addonMarker(appData)).toBe(beforeMarker);
    expect(readFileSync(path.join(js, "publish.xml"))).toEqual(beforePublish);
    expect(readFileSync(path.join(js, "wengge-excel-ai-addin-install-state.json"))).toEqual(
      beforeState,
    );
    expect(readFileSync(path.join(js, WPS_ADDON_DIRECTORY, "index.html"))).toEqual(beforeIndex);
    expect(activeTemps(appData)).toEqual([]);
    const st = statusWpsJsa({ appData, platform: "linux" });
    expect(st.current).toBe(true);
    expect(st.drift).toEqual([]);
  });

  it.each(["addon-move", "state-move", "publish-before", "publish-after"] as const)(
    "failAfter %s restores triple and stays current",
    (fp) => {
      const { packageDir } = buildPackage(`U_${fp}`);
      const appData = makeTempRoot("app-");
      installWpsJsa({ packageDir, appData, platform: "linux" });
      const js = jsaddonsOf(appData);
      const beforePublish = readFileSync(path.join(js, "publish.xml"));
      const beforeState = readFileSync(path.join(js, "wengge-excel-ai-addin-install-state.json"));
      const beforeIndex = readFileSync(path.join(js, WPS_ADDON_DIRECTORY, "index.html"));
      expect(() =>
        uninstallWpsJsa({ appData, platform: "linux", failAfter: fp }),
      ).toThrow(new RegExp(`failpoint:${fp}`));
      expect(readFileSync(path.join(js, "publish.xml"))).toEqual(beforePublish);
      expect(readFileSync(path.join(js, "wengge-excel-ai-addin-install-state.json"))).toEqual(
        beforeState,
      );
      expect(readFileSync(path.join(js, WPS_ADDON_DIRECTORY, "index.html"))).toEqual(beforeIndex);
      expect(activeTemps(appData)).toEqual([]);
      expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(true);
    },
  );

  it("successful uninstall removes own only; repeat leaves foreign publish bytes", () => {
    const { packageDir } = buildPackage("x");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`,
    );
    writeFileSync(path.join(js, "publish.xml.bak.2020"), "LEGACY");
    mkdirSync(path.join(js, "ExcelAIWps_0.1.30"), { recursive: true });
    writeFileSync(path.join(js, "ExcelAIWps_0.1.30/keep.txt"), "k");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    uninstallWpsJsa({ appData, platform: "linux" });
    expect(existsSync(path.join(js, WPS_ADDON_DIRECTORY))).toBe(false);
    const bytes = readFileSync(path.join(js, "publish.xml"));
    expect(
      parseJspluginsDocument(bytes.toString("utf8")).plugins.map(
        (p) => (p.attrs as Record<string, string>).name,
      ),
    ).toEqual(["ExcelAIWps"]);
    const un2 = uninstallWpsJsa({ appData, platform: "linux" });
    expect(un2.removed).toBe(false);
    expect(readFileSync(path.join(js, "publish.xml"))).toEqual(bytes);
    expect(readFileSync(path.join(js, "publish.xml.bak.2020"), "utf8")).toBe("LEGACY");
  });
});

describe("status honesty for own attrs", () => {
  it("missing debug / extra attr is not current", () => {
    const { packageDir } = buildPackage("s");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    const pub = path.join(js, "publish.xml");
    writeFileSync(
      pub,
      `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="${WPS_ADDON_NAME}" type="et" url="${WPS_PUBLISH_URL}" enable="enable_dev" />
</jsplugins>
`,
    );
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(false);

    installWpsJsa({ packageDir, appData, platform: "linux" });
    writeFileSync(
      pub,
      `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="${WPS_ADDON_NAME}" type="et" url="${WPS_PUBLISH_URL}" debug="" enable="enable_dev" extra="1" />
</jsplugins>
`,
    );
    expect(statusWpsJsa({ appData, platform: "linux" }).current).toBe(false);
  });
});

describe("publish backup rotation by mtime", () => {
  it("keeps newest 10 own backups; never touches publish.xml.bak.*", () => {
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    writeFileSync(path.join(js, "publish.xml.bak.2020"), "LEGACY");
    const base = Date.now() - 60_000;
    for (let i = 0; i < 12; i += 1) {
      const name = `${PUBLISH_BACKUP_PREFIX}ordered-${String(i).padStart(2, "0")}`;
      const full = path.join(js, name);
      writeFileSync(full, `b${i}`);
      const t = (base + i * 1000) / 1000;
      utimesSync(full, t, t);
    }
    rotateOwnPublishBackups(js);
    const own = readdirSync(js)
      .filter((n) => n.startsWith(PUBLISH_BACKUP_PREFIX))
      .sort();
    expect(own).toHaveLength(10);
    expect(own[0]).toContain("ordered-02");
    expect(own[9]).toContain("ordered-11");
    expect(readFileSync(path.join(js, "publish.xml.bak.2020"), "utf8")).toBe("LEGACY");
  });
});

describe("symlink fail-closed (independent appData per case)", () => {
  it("packageDir symlink refused", () => {
    const { packageDir } = buildPackage("p");
    const appData = makeTempRoot("app-");
    const link = path.join(makeTempRoot(), "pkg-link");
    if (trySymlink(packageDir, link) === "skip") return;
    expect(() =>
      installWpsJsa({ packageDir: link, appData, platform: "linux" }),
    ).toThrow(/symlink/i);
  });

  it("publish symlink refused", () => {
    const { packageDir } = buildPackage("p");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    const pub = path.join(js, "publish.xml");
    const real = readFileSync(pub);
    rmSync(pub);
    const target = path.join(js, "publish-real.xml");
    writeFileSync(target, real);
    if (trySymlink(target, pub) === "skip") return;
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow(/symlink/i);
  });

  it("addon dir symlink refused", () => {
    const { packageDir } = buildPackage("p");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    const addon = path.join(js, WPS_ADDON_DIRECTORY);
    const moved = path.join(js, "addon-real");
    // move real addon aside then symlink
    // use rename
    renameSync(addon, moved);
    if (trySymlink(moved, addon) === "skip") return;
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow(/symlink/i);
  });

  it("state symlink refused", () => {
    const { packageDir } = buildPackage("p");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    const state = path.join(js, "wengge-excel-ai-addin-install-state.json");
    const realBytes = readFileSync(state);
    rmSync(state);
    const target = path.join(js, "state-real.json");
    writeFileSync(target, realBytes);
    if (trySymlink(target, state) === "skip") return;
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow(/symlink/i);
  });

  it("dangling publish symlink refused", () => {
    const { packageDir } = buildPackage("p");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    const pub = path.join(js, "publish.xml");
    if (trySymlink(path.join(js, "missing-target.xml"), pub) === "skip") return;
    expect(() => installWpsJsa({ packageDir, appData, platform: "linux" })).toThrow(/symlink/i);
  });
});

describe("CLI process exit codes", () => {
  const node = process.execPath;
  const installCli = path.join(root, "scripts/wps-jsa-install.mjs");
  const statusCli = path.join(root, "scripts/wps-jsa-status.mjs");
  const uninstallCli = path.join(root, "scripts/wps-jsa-uninstall.mjs");

  it("help => 0", () => {
    for (const cli of [installCli, statusCli, uninstallCli]) {
      const r = spawnSync(node, [cli, "--help"], { encoding: "utf8" });
      expect(r.status, cli).toBe(0);
    }
  });

  it("missing value / option-as-value / unknown => 1", () => {
    expect(spawnSync(node, [installCli, "--package-dir"], { encoding: "utf8" }).status).toBe(1);
    expect(
      spawnSync(node, [installCli, "--app-data", "--help"], { encoding: "utf8" }).status,
    ).toBe(1);
    expect(spawnSync(node, [statusCli, "--nope"], { encoding: "utf8" }).status).toBe(1);
  });

  it("status empty appData => 2 JSON current false; uninstall => 0", () => {
    const appData = makeTempRoot("cli-app-");
    const st = spawnSync(node, [statusCli, "--app-data", appData], { encoding: "utf8" });
    expect(st.status).toBe(2);
    const json = JSON.parse(st.stdout);
    expect(json.current).toBe(false);
    const un = spawnSync(node, [uninstallCli, "--app-data", appData], { encoding: "utf8" });
    expect(un.status).toBe(0);
  });

  it("parser rejects package-dir + git-sha", () => {
    expect(() =>
      parseWpsInstallCliArgs(["--package-dir", "./dist", "--git-sha", "abc"], {
        allowGitSha: true,
        allowPackageDir: true,
      }),
    ).toThrow(/git-sha/);
    expect(() => resolveAppDataRoot({ platform: "linux", env: {} })).toThrow(/--app-data/);
  });
});

describe("package validate", () => {
  it("accepts built package", () => {
    const { packageDir } = buildPackage("v");
    expect(validateWpsPackageDir(packageDir).buildInfo.target).toBe("wps-jsa");
  });
});

describe("Phase57 dry-run zero AppData writes", () => {
  function snapshotTree(rootDir: string): { entries: string[]; files: Record<string, { size: number; mtimeMs: number; sha: string }> } {
    const entries: string[] = [];
    const files: Record<string, { size: number; mtimeMs: number; sha: string }> = {};
    if (!existsSync(rootDir)) return { entries, files };
    const walk = (dir: string, rel = "") => {
      for (const name of readdirSync(dir).sort()) {
        const abs = path.join(dir, name);
        const r = rel ? `${rel}/${name}` : name;
        entries.push(r);
        const st = statSync(abs);
        if (st.isDirectory()) walk(abs, r);
        else if (st.isFile()) {
          files[r] = {
            size: st.size,
            mtimeMs: st.mtimeMs,
            sha: createHash("sha256").update(readFileSync(abs)).digest("hex"),
          };
        }
      }
    };
    walk(rootDir);
    return { entries, files };
  }

  it("absent appData: succeeds and path still missing", () => {
    const { packageDir } = buildPackage("dry-absent");
    const appData = path.join(makeTempRoot("parent-"), "no-appdata-yet");
    expect(existsSync(appData)).toBe(false);
    const r = installWpsJsa({
      packageDir,
      appData,
      platform: "linux",
      dryRun: true,
      createWpsPackage: () => {
        throw new Error("createWpsPackage must not run with packageDir");
      },
    });
    expect(r.dryRun).toBe(true);
    expect(r.wouldCreateJsaddons).toBe(true);
    expect(r.wouldCreatePublish).toBe(true);
    if (!("wouldWriteState" in r) || !("wouldInstall" in r)) {
      throw new Error("expected dry-run plan fields");
    }
    expect(r.wouldWriteState).toBe(true);
    expect(r.wouldInstall).toBe(true);
    expect(existsSync(appData)).toBe(false);
  });

  it("empty existing appData/jsaddons: bytes+mtime+entries unchanged", () => {
    const { packageDir } = buildPackage("dry-empty");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    const before = snapshotTree(appData);
    const r = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    expect(r.dryRun).toBe(true);
    expect(r.wouldCreateJsaddons).toBe(false);
    expect(r.wouldCreatePublish).toBe(true);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("real host fixture preserved; bytes/mtime unchanged", () => {
    const { packageDir } = buildPackage("dry-host");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(path.join(js, "ExcelAIWps_0.1.30"), { recursive: true });
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`,
    );
    writeFileSync(path.join(js, "publish.xml.bak.2020"), "LEGACY");
    writeFileSync(path.join(js, "ExcelAIWps_0.1.30/keep.txt"), "k");
    const before = snapshotTree(appData);
    const r = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    expect(r.preservedPluginNames).toContain("ExcelAIWps");
    expect(r.warnings.some((w: string) => /ExcelAIWps/i.test(w))).toBe(true);
    expect(r.wouldUpdatePublish).toBe(true);
    expect(r.existingOwnEntry).toBe(false);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("already installed current: no backup/temp/state mutation", () => {
    const { packageDir } = buildPackage("dry-cur");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    const statePath = path.join(js, "wengge-excel-ai-addin-install-state.json");
    const stateBefore = readFileSync(statePath, "utf8");
    const before = snapshotTree(appData);
    const r = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    expect(r.existingOwnEntry).toBe(true);
    expect(r.wouldReplaceAddon).toBe(true);
    expect(r.wouldUpdatePublish).toBe(false);
    expect(snapshotTree(appData)).toEqual(before);
    expect(readFileSync(statePath, "utf8")).toBe(stateBefore);
    expect(activeTemps(appData)).toEqual([]);
  });

  it("own attrs drift: wouldUpdatePublish true, no write", () => {
    const { packageDir } = buildPackage("dry-drift");
    const appData = makeTempRoot("app-");
    installWpsJsa({ packageDir, appData, platform: "linux" });
    const js = jsaddonsOf(appData);
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0" encoding="UTF-8"?>
<jsplugins>
  <jsplugin name="${WPS_ADDON_NAME}" type="et" url="${WPS_PUBLISH_URL}" enable="enable_dev" />
</jsplugins>
`,
    );
    const before = snapshotTree(appData);
    const r = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    expect(r.wouldUpdatePublish).toBe(true);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("malformed publish / DOCTYPE fail closed with zero tree change", () => {
    const { packageDir } = buildPackage("dry-bad");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY y "z">]><jsplugins></jsplugins>`,
    );
    const before = snapshotTree(appData);
    expect(() =>
      installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true }),
    ).toThrow(/DOCTYPE|ENTITY/i);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("package-dir dry-run never calls createWpsPackage", () => {
    const { packageDir } = buildPackage("dry-nobuild");
    const appData = makeTempRoot("app-");
    let called = 0;
    installWpsJsa({
      packageDir,
      appData,
      platform: "linux",
      dryRun: true,
      createWpsPackage: () => {
        called += 1;
        throw new Error("build");
      },
    });
    expect(called).toBe(0);
  });

  it("plan fields match subsequent real install surface", () => {
    const { packageDir } = buildPackage("plan-real");
    const appData = makeTempRoot("app-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
</jsplugins>
`,
    );
    const plan = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    expect(plan.preservedPluginNames).toEqual(["ExcelAIWps"]);
    expect(plan.wouldCreateJsaddons).toBe(false);
    expect(plan.wouldReplaceAddon).toBe(false);
    const real = installWpsJsa({ packageDir, appData, platform: "linux" });
    expect(real.ok).toBe(true);
    expect(real.preservedPluginNames).toEqual(plan.preservedPluginNames);
    const st = statusWpsJsa({ appData, platform: "linux" });
    expect(st.current).toBe(true);
    const names = parseJspluginsDocument(
      readFileSync(path.join(js, "publish.xml"), "utf8"),
    ).plugins.map((p) => (p.attrs as Record<string, string>).name);
    expect(names).toContain("ExcelAIWps");
    expect(names).toContain(WPS_ADDON_NAME);
  });

  it("CLI: dry-run exit 0 JSON; status/uninstall --dry-run exit 1", () => {
    const { packageDir } = buildPackage("cli-dry");
    const appData = makeTempRoot("app-");
    const node = process.execPath;
    const installCli = path.join(root, "scripts/wps-jsa-install.mjs");
    const statusCli = path.join(root, "scripts/wps-jsa-status.mjs");
    const uninstallCli = path.join(root, "scripts/wps-jsa-uninstall.mjs");
    const dry = spawnSync(
      node,
      [installCli, "--package-dir", packageDir, "--app-data", appData, "--dry-run"],
      { encoding: "utf8" },
    );
    expect(dry.status).toBe(0);
    const json = JSON.parse(dry.stdout);
    expect(json.dryRun).toBe(true);
    expect(json.wouldInstall).toBe(true);
    expect(existsSync(jsaddonsOf(appData))).toBe(false);
    expect(spawnSync(node, [statusCli, "--dry-run"], { encoding: "utf8" }).status).toBe(1);
    expect(spawnSync(node, [uninstallCli, "--dry-run"], { encoding: "utf8" }).status).toBe(1);
  });
});


describe("Phase57.1 plan ancestry + public name projection", () => {
  it("dry-run fails closed when appData root is a symlink (no tree writes)", () => {
    const { packageDir } = buildPackage("anc-root");
    const base = makeTempRoot("anc-");
    const real = path.join(base, "real-appdata");
    const link = path.join(base, "link-appdata");
    mkdirSync(real, { recursive: true });
    try {
      symlinkSync(real, link);
    } catch (error) {
      const code = error && typeof error === "object" && "code" in error ? String((error as { code?: string }).code) : "";
      if (code === "EPERM" || code === "ENOTSUP") return;
      throw error;
    }
    const before = snapshotTree(base);
    expect(() =>
      installWpsJsa({ packageDir, appData: link, platform: "linux", dryRun: true }),
    ).toThrow(/symlink|junction|directory/i);
    expect(snapshotTree(base)).toEqual(before);
  });

  it("dry-run fails when intermediate kingsoft is non-directory; matches real install", () => {
    const { packageDir } = buildPackage("anc-file");
    const appData = makeTempRoot("ancf-");
    writeFileSync(path.join(appData, "kingsoft"), "not-a-dir");
    const before = snapshotTree(appData);
    expect(() =>
      installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true }),
    ).toThrow(/non-directory|symlink|directory/i);
    expect(() =>
      installWpsJsa({ packageDir, appData, platform: "linux" }),
    ).toThrow(/non-directory|symlink|directory/i);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("dry-run fails on dangling intermediate symlink with zero writes", () => {
    const { packageDir } = buildPackage("anc-dang");
    const appData = makeTempRoot("ancd-");
    mkdirSync(path.join(appData, "kingsoft"), { recursive: true });
    try {
      symlinkSync(path.join(appData, "missing-target"), path.join(appData, "kingsoft", "wps"));
    } catch (error) {
      const code = error && typeof error === "object" && "code" in error ? String((error as { code?: string }).code) : "";
      if (code === "EPERM" || code === "ENOTSUP") return;
      throw error;
    }
    const before = snapshotTree(appData);
    expect(() =>
      installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true }),
    ).toThrow(/symlink|junction/i);
    expect(snapshotTree(appData)).toEqual(before);
  });

  it("unsafe foreign plugin name is preserved in publish merge but not in dry-run JSON", () => {
    const { packageDir } = buildPackage("name-secret");
    const appData = makeTempRoot("sec-");
    const js = jsaddonsOf(appData);
    mkdirSync(js, { recursive: true });
    // Attribute parser rejects <>; use long control-char name without <>.
    const evilName = `Evil\tPlugin_SECRETTOKEN_do_not_echo_${"x".repeat(80)}`;
    writeFileSync(
      path.join(js, "publish.xml"),
      `<?xml version="1.0" encoding="utf-8"?>
<jsplugins>
  <jsplugin name="ExcelAIWps" enable="enable_dev" url="file://" type="et" version="0.1.30" />
  <jsplugin name="${evilName}" enable="enable_dev" url="file://" type="et" />
</jsplugins>
`,
    );
    const before = readFileSync(path.join(js, "publish.xml"), "utf8");
    const dry = installWpsJsa({ packageDir, appData, platform: "linux", dryRun: true });
    const encoded = JSON.stringify(dry);
    expect(encoded).not.toContain("SECRETTOKEN");
    expect(encoded).not.toContain("api_key_");
    expect(encoded).not.toContain(evilName);
    expect(dry.preservedPluginNames).toContain("ExcelAIWps");
    expect(dry.preservedPluginNames).toContain("(unsafe-plugin-name)");
    expect(dry.warnings.every((w: string) => !w.includes("SECRETTOKEN"))).toBe(true);
    // dry-run zero write
    expect(readFileSync(path.join(js, "publish.xml"), "utf8")).toBe(before);

    const real = installWpsJsa({ packageDir, appData, platform: "linux" });
    const pub = readFileSync(path.join(js, "publish.xml"), "utf8");
    expect(pub).toContain(evilName);
    expect(pub).toContain("ExcelAIWps");
    expect(pub).toContain(WPS_ADDON_NAME);
    expect(JSON.stringify(real)).not.toContain("SECRETTOKEN");
    expect(real.preservedPluginNames).toContain("(unsafe-plugin-name)");
  });
});
