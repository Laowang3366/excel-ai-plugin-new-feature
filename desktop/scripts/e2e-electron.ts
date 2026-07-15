/**
 * Minimal Electron E2E via playwright-core _electron (no browser download).
 * Covers: remote navigation block + openExternal, tool approval UI, settings IPC,
 * update state projection, stream delta order across interrupt/resume.
 */
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { createRequire } from "node:module";

import { _electron as electron, type ElectronApplication, type Page } from "playwright-core";

import {
  ensureChatWorkspace,
  scenarioNavigationAndExternal,
  scenarioSettingsIpc,
  scenarioStreamOrder,
  scenarioToolApproval,
  scenarioUpdateState,
  stubOpenExternal,
  type ScenarioResult,
} from "./e2e-electron-scenarios";

const require = createRequire(path.join(process.cwd(), "package.json"));
const desktopRoot = process.cwd();
const mainPath = path.join(desktopRoot, "dist-electron", "main.js");
const electronBinary = require("electron") as string;

async function main(): Promise<void> {
  if (!existsSync(mainPath)) {
    throw new Error(
      `Electron E2E requires built main at ${mainPath}. Run npm run build first (not electron:build).`,
    );
  }

  const userDataDir = await mkdtemp(path.join(os.tmpdir(), "wengge-e2e-userdata-"));
  let app: ElectronApplication | undefined;
  const results: ScenarioResult[] = [];

  try {
    app = await electron.launch({
      executablePath: electronBinary,
      args: [mainPath, `--user-data-dir=${userDataDir}`],
      cwd: desktopRoot,
      env: {
        ...process.env,
        ELECTRON_DISABLE_SECURITY_WARNINGS: "1",
        WENGE_UPDATE_BASE_URL: "http://127.0.0.1:9",
      },
      timeout: 90_000,
    });

    const page = await waitForMainWindow(app);
    await page.waitForLoadState("domcontentloaded");
    await page.waitForSelector("#root", { timeout: 60_000 });
    await ensureChatWorkspace(page);

    await stubOpenExternal(app);
    results.push(await scenarioNavigationAndExternal(app, page));
    // Blocked remote navigation can leave Playwright waiting on an in-flight navigation.
    await page.reload({ waitUntil: "domcontentloaded" });
    await ensureChatWorkspace(page);
    await stubOpenExternal(app);
    results.push(await scenarioToolApproval(app, page));
    results.push(await scenarioSettingsIpc(page));
    results.push(await scenarioUpdateState(app, page));
    results.push(await scenarioStreamOrder(app, page));

    const failed = results.filter((item) => !item.ok);
    process.stdout.write(`${JSON.stringify({ ok: failed.length === 0, results }, null, 2)}\n`);
    if (failed.length > 0) process.exitCode = 1;
  } finally {
    if (app) {
      await app.close().catch(() => undefined);
    }
    await rm(userDataDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

async function waitForMainWindow(app: ElectronApplication): Promise<Page> {
  const existing = app.windows();
  if (existing[0]) return existing[0];
  return app.firstWindow();
}

void main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? error.stack || error.message : String(error)}\n`,
  );
  process.exitCode = 1;
});
