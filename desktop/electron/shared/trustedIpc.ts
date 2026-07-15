import {
  ipcMain,
  type BrowserWindow,
  type IpcMainEvent,
  type IpcMainInvokeEvent,
} from "electron";

import { createIpcRateLimiter } from "./ipcRateLimiter";

type MainWindowRef = () => BrowserWindow | null;

let mainWindowRef: MainWindowRef | null = null;
const rateLimiter = createIpcRateLimiter();

export function configureTrustedIpcSender(ref: MainWindowRef): void {
  mainWindowRef = ref;
}

export function isTrustedRendererUrl(
  url: string,
  devServerUrl = process.env.VITE_DEV_SERVER_URL
): boolean {
  try {
    const candidate = new URL(url);
    if (!devServerUrl) return candidate.protocol === "file:";

    const developmentOrigin = new URL(devServerUrl);
    return (
      (candidate.protocol === "http:" || candidate.protocol === "https:") &&
      candidate.origin === developmentOrigin.origin
    );
  } catch {
    return false;
  }
}

export function isTrustedIpcSender(
  event: Pick<IpcMainInvokeEvent, "sender" | "senderFrame">,
  getMainWindow: MainWindowRef | null = mainWindowRef,
  devServerUrl = process.env.VITE_DEV_SERVER_URL
): boolean {
  const mainWindow = getMainWindow?.();
  if (!mainWindow || mainWindow.isDestroyed() || mainWindow.webContents.isDestroyed()) {
    return false;
  }

  const senderFrame = event.senderFrame;
  return (
    event.sender === mainWindow.webContents &&
    senderFrame === mainWindow.webContents.mainFrame &&
    isTrustedRendererUrl(senderFrame.url, devServerUrl)
  );
}

export function assertTrustedIpcSender(
  event: Pick<IpcMainInvokeEvent, "sender" | "senderFrame">
): void {
  if (!isTrustedIpcSender(event)) {
    throw new Error("unauthorized_ipc_sender");
  }
}

export const trustedIpcMain = {
  handle<TArgs extends unknown[], TResult>(
    channel: string,
    listener: (event: IpcMainInvokeEvent, ...args: TArgs) => TResult
  ): void {
    ipcMain.handle(channel, (event, ...args) => {
      assertTrustedIpcSender(event);
      rateLimiter.assertAllowed(channel, event);
      return listener(event, ...(args as TArgs));
    });
  },
  on<TArgs extends unknown[]>(
    channel: string,
    listener: (event: IpcMainEvent, ...args: TArgs) => void
  ): void {
    ipcMain.on(channel, (event, ...args) => {
      assertTrustedIpcSender(event);
      listener(event, ...(args as TArgs));
    });
  },
  removeHandler(channel: string): void {
    ipcMain.removeHandler(channel);
  },
};
