import type { ElectronApplication, Page } from "playwright-core";

import type { ScenarioResult } from "./e2e-electron-scenarios";

/** Markdown openExternal + blocked window.open + blocked location= (order avoids Playwright hang). */
export async function scenarioNavigationAndExternal(
  app: ElectronApplication,
  page: Page,
): Promise<ScenarioResult> {
  const name = "navigation-external";
  const markdownUrl = "https://example.com/c01-md-link";
  try {
    await app.evaluate(() => {
      const bucket = (globalThis as { __e2eOpenedUrls?: string[] }).__e2eOpenedUrls;
      if (bucket) bucket.length = 0;
    });
    await sendAgentEvent(app, { type: "turn_started", turnId: "e2e-c01-md" });
    await sendAgentEvent(app, {
      type: "stream_delta",
      delta: `See [C01 docs](${markdownUrl}) for details.`,
      itemType: "assistant_message",
      roundId: 1,
    });
    const link = page.locator(`a[href="${markdownUrl}"]`);
    await link.waitFor({ state: "attached", timeout: 10_000 });
    const urlBeforeClick = page.url();
    await link.click({ noWaitAfter: true });
    await page.waitForTimeout(400);
    if (page.url() !== urlBeforeClick) {
      return { name, ok: false, detail: `Markdown click navigated: ${page.url()}` };
    }
    const opened = await app.evaluate(
      () => (globalThis as { __e2eOpenedUrls?: string[] }).__e2eOpenedUrls || [],
    );
    if (!opened.some((url) => url === markdownUrl || url.includes(markdownUrl))) {
      return {
        name,
        ok: false,
        detail: `Markdown openExternal missing: ${JSON.stringify(opened)}`,
      };
    }

    const windowsBefore = await countWindows(app);
    const urlBeforeOpen = page.url();
    await page.evaluate(() => {
      window.open("https://evil.example/popup", "_blank");
    });
    await page.waitForTimeout(400);
    if ((await countWindows(app)) > windowsBefore) {
      return { name, ok: false, detail: `window.open created BrowserWindow` };
    }
    if (page.url() !== urlBeforeOpen) {
      return { name, ok: false, detail: `window.open changed url: ${page.url()}` };
    }

    const beforeUrl = page.url();
    await page.evaluate(() => {
      window.location.href = "https://evil.example/phish";
    });
    await page.waitForTimeout(400);
    if (page.url() !== beforeUrl) {
      return { name, ok: false, detail: `location not blocked: ${page.url()}` };
    }

    return {
      name,
      ok: true,
      detail: `Markdown click openExternal=${markdownUrl}; blocked location; blocked window.open`,
    };
  } catch (error) {
    return {
      name,
      ok: false,
      detail: error instanceof Error ? error.stack || error.message : String(error),
    };
  }
}

async function countWindows(app: ElectronApplication): Promise<number> {
  return app.evaluate(
    ({ BrowserWindow }) => BrowserWindow.getAllWindows().filter((win) => !win.isDestroyed()).length,
  );
}

async function sendAgentEvent(
  app: ElectronApplication,
  event: Record<string, unknown>,
): Promise<void> {
  await app.evaluate(({ BrowserWindow }, payload) => {
    const win = BrowserWindow.getAllWindows()[0];
    if (!win || win.isDestroyed()) throw new Error("no main window");
    win.webContents.send("agent:event", payload);
  }, event);
}
