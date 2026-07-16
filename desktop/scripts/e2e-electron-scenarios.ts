import type { ElectronApplication, Page } from "playwright-core";

export type ScenarioResult = { name: string; ok: boolean; detail?: string };

/** Fresh profiles open Settings when unconfigured; ToolConfirm lives on ChatPage. */
export async function ensureChatWorkspace(page: Page): Promise<void> {
  await page.waitForFunction(() => !document.querySelector(".app-loading"), null, {
    timeout: 60_000,
  });

  await page.evaluate(async () => {
    const activeId = "e2e-provider";
    await window.electronAPI.settings.set("aiProviders", {
      [activeId]: {
        id: activeId,
        name: "E2E Provider",
        provider: "openai",
        apiKey: "e2e-dummy-key",
        baseUrl: "https://example.com/v1",
        model: "e2e-model",
        defaultModel: "e2e-model",
      },
    });
    await window.electronAPI.settings.set("activeProvider", activeId);
  });

  await page.reload({ waitUntil: "domcontentloaded" });
  await page.waitForFunction(() => !document.querySelector(".app-loading"), null, {
    timeout: 60_000,
  });

  const back = page.locator(".settings-back-btn");
  if (await back.count()) {
    await back.first().click();
    await page.waitForTimeout(300);
  }

  await page.waitForFunction(
    () =>
      Boolean(
        document.querySelector(".composer-area") ||
        document.querySelector(".chat-page") ||
        document.querySelector(".welcome-panel"),
      ),
    null,
    { timeout: 30_000 },
  );
}

export async function stubOpenExternal(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ shell }) => {
    const bucket = ((globalThis as { __e2eOpenedUrls?: string[] }).__e2eOpenedUrls = []);
    const original = shell.openExternal.bind(shell);
    (shell as { openExternal: (url: string) => Promise<void> }).openExternal = async (
      url: string,
    ) => {
      bucket.push(String(url));
    };
    (globalThis as { __e2eRestoreOpenExternal?: () => void }).__e2eRestoreOpenExternal = () => {
      (shell as { openExternal: typeof original }).openExternal = original;
    };
  });
}

export async function scenarioToolApproval(
  app: ElectronApplication,
  page: Page,
): Promise<ScenarioResult> {
  const name = "tool-approval";
  try {
    await app.evaluate(({ ipcMain }) => {
      (
        globalThis as { __e2eToolConfirms?: Array<{ toolCallId: string; alwaysAllow?: boolean }> }
      ).__e2eToolConfirms = [];
      ipcMain.removeHandler("tool:confirm");
      ipcMain.handle("tool:confirm", (_event, toolCallId: unknown, alwaysAllow?: unknown) => {
        (
          globalThis as {
            __e2eToolConfirms?: Array<{ toolCallId: string; alwaysAllow?: boolean }>;
          }
        ).__e2eToolConfirms?.push({
          toolCallId: String(toolCallId),
          alwaysAllow: alwaysAllow === true,
        });
      });
    });

    await app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0];
      if (!win || win.isDestroyed()) throw new Error("no main window");
      win.webContents.send("agent:event", {
        type: "tool_approval_required",
        toolCallId: "e2e-tool-1",
        toolName: "range.write",
        arguments: { sheetName: "Sheet1", range: "A1" },
        riskLevel: "moderate",
        description: "E2E approval probe",
        canAlwaysAllow: false,
      });
    });

    await page.locator(".tool-confirm-dialog").waitFor({ state: "attached", timeout: 10_000 });
    const toolName = await page.locator(".tool-confirm-name").innerText();
    if (!toolName.includes("range.write")) {
      return { name, ok: false, detail: `unexpected tool name: ${toolName}` };
    }
    await page.locator(".tool-confirm-approve").click({ noWaitAfter: true });
    await page.locator(".tool-confirm-dialog").waitFor({ state: "detached", timeout: 10_000 });
    const confirms = await app.evaluate(() => {
      return (globalThis as { __e2eToolConfirms?: Array<{ toolCallId: string }> })
        .__e2eToolConfirms;
    });
    if (!confirms?.some((item) => item.toolCallId === "e2e-tool-1")) {
      return { name, ok: false, detail: `confirm not returned: ${JSON.stringify(confirms)}` };
    }
    return { name, ok: true };
  } catch (error) {
    return { name, ok: false, detail: errorMessage(error) };
  }
}

export async function scenarioSettingsIpc(page: Page): Promise<ScenarioResult> {
  const name = "settings-ipc";
  try {
    const result = await page.evaluate(async () => {
      const previous = String(
        (await window.electronAPI.settings.get("permissionMode")) || "normal",
      );
      // Target must differ from previous so a no-op set cannot pass.
      const target = previous === "confirm_all" ? "normal" : "confirm_all";
      await window.electronAPI.settings.set("permissionMode", target);
      const next = String(await window.electronAPI.settings.get("permissionMode"));
      await window.electronAPI.settings.set("permissionMode", previous);
      const restored = String(await window.electronAPI.settings.get("permissionMode"));
      return { previous, target, next, restored };
    });
    if (result.next !== result.target || result.next === result.previous) {
      return {
        name,
        ok: false,
        detail: `set did not change value: ${JSON.stringify(result)}`,
      };
    }
    if (result.restored !== result.previous) {
      return {
        name,
        ok: false,
        detail: `restore failed: ${JSON.stringify(result)}`,
      };
    }
    return { name, ok: true, detail: JSON.stringify(result) };
  } catch (error) {
    return { name, ok: false, detail: errorMessage(error) };
  }
}

export async function scenarioUpdateState(
  app: ElectronApplication,
  page: Page,
): Promise<ScenarioResult> {
  const name = "update-state";
  try {
    await openUpdateSettingsUi(page);

    await app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows()[0];
      if (!win || win.isDestroyed()) throw new Error("no main window");
      win.webContents.send("update:stateChanged", {
        phase: "available",
        currentVersion: "0.0.0",
        availableVersion: "9.9.9-e2e",
        installerAvailable: true,
        hotPatchAvailable: false,
        releaseNotes: ["e2e"],
      });
    });

    await page.locator(".update-settings").waitFor({ state: "attached", timeout: 10_000 });
    await page.waitForFunction(
      () => {
        const root = document.querySelector(".update-settings");
        if (!root) return false;
        const text = root.textContent || "";
        return (
          text.includes("9.9.9-e2e") && Boolean(document.querySelector(".update-status-available"))
        );
      },
      null,
      { timeout: 10_000 },
    );
    const uiText = await page.locator(".update-settings").innerText();
    if (!uiText.includes("9.9.9-e2e")) {
      return { name, ok: false, detail: `available version not rendered: ${uiText.slice(0, 300)}` };
    }
    const back = page.locator(".settings-back-btn");
    if (await back.count()) {
      await back.first().click({ noWaitAfter: true });
      await page.waitForTimeout(300);
    }
    return { name, ok: true, detail: "UI shows 9.9.9-e2e + update-status-available" };
  } catch (error) {
    return { name, ok: false, detail: errorMessage(error) };
  }
}

async function openUpdateSettingsUi(page: Page): Promise<void> {
  const settingsTrigger = page.locator(".sidebar-nav-btn").filter({ hasText: /设置|Settings/i });
  if (await settingsTrigger.count()) {
    await settingsTrigger.first().click({ noWaitAfter: true });
    await page.waitForTimeout(200);
    const general = page
      .locator(".sidebar-settings-menu-item")
      .filter({ hasText: /设置|Settings/i });
    if (await general.count()) {
      await general.first().click({ noWaitAfter: true });
      await page.waitForTimeout(300);
    }
  }
  await page.locator(".settings-shell").waitFor({ state: "attached", timeout: 15_000 });
  const updatesNav = page
    .locator(".settings-sidebar-item")
    .filter({ hasText: /软件更新|Updates/i });
  await updatesNav.first().click({ noWaitAfter: true });
  await page.locator(".update-settings").waitFor({ state: "attached", timeout: 10_000 });
}

export async function scenarioStreamOrder(
  app: ElectronApplication,
  page: Page,
): Promise<ScenarioResult> {
  const name = "stream-order";
  try {
    await sendAgentEvent(app, { type: "turn_started", turnId: "e2e-turn-1" });
    await sendAgentEvent(app, {
      type: "stream_delta",
      delta: "Alpha",
      itemType: "assistant_message",
      roundId: 1,
    });
    await sendAgentEvent(app, {
      type: "stream_delta",
      delta: "-Beta",
      itemType: "assistant_message",
      roundId: 1,
    });
    await page.waitForTimeout(120);
    const mid = await page.evaluate(() => document.body.innerText);
    if (!mid.includes("Alpha-Beta")) {
      return {
        name,
        ok: false,
        detail: `pre-interrupt missing exact Alpha-Beta; body=${mid.slice(0, 400)}`,
      };
    }
    if (mid.includes("Beta-Alpha")) {
      return { name, ok: false, detail: "pre-interrupt stream order inverted" };
    }

    await sendAgentEvent(app, {
      type: "item_started",
      item: {
        type: "tool_call",
        id: "e2e-freeze-tool",
        toolName: "e2e.freeze",
        arguments: {},
        status: "running",
        timestamp: Date.now(),
      },
    });
    await page.waitForTimeout(80);

    await sendAgentEvent(app, { type: "turn_interrupted" });
    await sendAgentEvent(app, { type: "turn_started", turnId: "e2e-turn-2" });
    await sendAgentEvent(app, {
      type: "stream_delta",
      delta: "Gamma",
      itemType: "assistant_message",
      roundId: 2,
    });
    await sendAgentEvent(app, {
      type: "stream_delta",
      delta: "-Delta",
      itemType: "assistant_message",
      roundId: 2,
    });
    await page.waitForTimeout(150);

    const body = await page.evaluate(() => document.body.innerText);
    if (!body.includes("Alpha-Beta")) {
      return {
        name,
        ok: false,
        detail: `post-resume lost pre-interrupt Alpha-Beta; body=${body.slice(0, 400)}`,
      };
    }
    if (!body.includes("Gamma-Delta")) {
      return {
        name,
        ok: false,
        detail: `post-resume missing exact Gamma-Delta; body=${body.slice(0, 400)}`,
      };
    }
    if (body.includes("Delta-Gamma") || body.includes("Beta-Alpha")) {
      return { name, ok: false, detail: "stream order inverted after resume" };
    }
    const alphaIdx = body.indexOf("Alpha-Beta");
    const gammaIdx = body.indexOf("Gamma-Delta");
    if (alphaIdx < 0 || gammaIdx < 0 || alphaIdx > gammaIdx) {
      return {
        name,
        ok: false,
        detail: `sequence order wrong: alpha@${alphaIdx} gamma@${gammaIdx}`,
      };
    }
    return { name, ok: true };
  } catch (error) {
    return { name, ok: false, detail: errorMessage(error) };
  }
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.stack || error.message : String(error);
}
