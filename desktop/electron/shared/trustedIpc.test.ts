import { describe, expect, it, vi } from "vitest";

import { isTrustedIpcSender, isTrustedRendererUrl } from "./trustedIpc";

function createSender(url: string, topLevel = true) {
  const mainFrame = { url };
  const webContents = {
    isDestroyed: vi.fn(() => false),
    mainFrame,
  };
  return {
    event: {
      sender: webContents,
      senderFrame: topLevel ? mainFrame : { url },
    },
    getMainWindow: () => ({
      isDestroyed: vi.fn(() => false),
      webContents,
    }),
  };
}

describe("trusted IPC sender policy", () => {
  it("accepts the production file renderer and exact development origin", () => {
    expect(isTrustedRendererUrl("file:///C:/app/dist/index.html", undefined)).toBe(true);
    expect(
      isTrustedRendererUrl("http://localhost:5173/chat", "http://localhost:5173")
    ).toBe(true);
  });

  it("rejects remote pages and protocol confusion", () => {
    expect(isTrustedRendererUrl("https://attacker.example", undefined)).toBe(false);
    expect(
      isTrustedRendererUrl("http://localhost.attacker.example:5173", "http://localhost:5173")
    ).toBe(false);
    expect(isTrustedRendererUrl("javascript:alert(1)", undefined)).toBe(false);
  });

  it("requires the current main window and its top-level frame", () => {
    const trusted = createSender("file:///C:/app/dist/index.html");
    expect(isTrustedIpcSender(trusted.event as never, trusted.getMainWindow as never, undefined)).toBe(true);

    const subframe = createSender("file:///C:/app/dist/index.html", false);
    expect(isTrustedIpcSender(subframe.event as never, subframe.getMainWindow as never, undefined)).toBe(false);

    const remote = createSender("https://attacker.example");
    expect(isTrustedIpcSender(remote.event as never, remote.getMainWindow as never, undefined)).toBe(false);
  });
});
