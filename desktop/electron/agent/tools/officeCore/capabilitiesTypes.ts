import type { OfficeActionApp, OfficeActionEngine } from "./types";

export interface OfficeCapability {
  app: OfficeActionApp;
  operation: string;
  preferredEngine: OfficeActionEngine;
  writesFile: boolean;
  fallback: "none" | "needsCom";
}
