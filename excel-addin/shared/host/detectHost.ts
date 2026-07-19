import { OfficeJsAdapter } from "./officeJsAdapter";
import type { HostAdapter, HostKind } from "./types";
import { WpsJsaAdapter } from "./wpsJsaAdapter";

export function detectHostKind(): HostKind {
  if (typeof window === "undefined") return "unknown";
  if (window.Excel?.run) return "office-js";
  if (window.Application?.ActiveWorkbook || window.Application?.Name) return "wps-jsa";
  if (window.Office) return "office-js";
  return "unknown";
}

export function createHostAdapter(kind: HostKind = detectHostKind()): HostAdapter {
  if (kind === "wps-jsa") return new WpsJsaAdapter();
  return new OfficeJsAdapter();
}

export async function waitForOfficeReady(timeoutMs = 5000): Promise<HostKind> {
  if (typeof window === "undefined") return "unknown";
  if (window.Excel?.run) return "office-js";
  if (window.Application) return "wps-jsa";

  return new Promise((resolve) => {
    let settled = false;
    const finish = (kind: HostKind) => {
      if (settled) return;
      settled = true;
      resolve(kind);
    };

    const timer = setTimeout(() => finish(detectHostKind()), timeoutMs);

    if (window.Office?.onReady) {
      window.Office.onReady(() => {
        clearTimeout(timer);
        finish(detectHostKind());
      });
      return;
    }

    clearTimeout(timer);
    finish(detectHostKind());
  });
}
