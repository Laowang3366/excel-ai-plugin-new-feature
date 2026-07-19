/**
 * Minimal ExcelApi 1.9 PageLayout.headersFooters.defaultForAllPages helpers.
 * Types live in officeJsRuntime.ts (single facade source).
 */
import type { ExcelHeaderFooter } from "./officeJsRuntime";

export const HEADER_FOOTER_SLOTS = [
  "leftHeader",
  "centerHeader",
  "rightHeader",
  "leftFooter",
  "centerFooter",
  "rightFooter",
] as const;

export type HeaderFooterSlot = (typeof HEADER_FOOTER_SLOTS)[number];

const HEADER_FOOTER_LOAD =
  "leftHeader,centerHeader,rightHeader,leftFooter,centerFooter,rightFooter";

export function requireHeaderFooterString(value: unknown, field: string): string {
  if (typeof value !== "string") {
    throw new Error(`PageLayout.headersFooters.defaultForAllPages.${field} is not a loaded string`);
  }
  return value;
}

export function readDefaultHeadersFooters(hf: ExcelHeaderFooter): {
  headers: { left: string; center: string; right: string };
  footers: { left: string; center: string; right: string };
} {
  for (const slot of HEADER_FOOTER_SLOTS) {
    if (!(slot in hf)) {
      throw new Error(
        `PageLayout.headersFooters.defaultForAllPages.${slot} is missing on host layout object`,
      );
    }
  }
  return {
    headers: {
      left: requireHeaderFooterString(hf.leftHeader, "leftHeader"),
      center: requireHeaderFooterString(hf.centerHeader, "centerHeader"),
      right: requireHeaderFooterString(hf.rightHeader, "rightHeader"),
    },
    footers: {
      left: requireHeaderFooterString(hf.leftFooter, "leftFooter"),
      center: requireHeaderFooterString(hf.centerFooter, "centerFooter"),
      right: requireHeaderFooterString(hf.rightFooter, "rightFooter"),
    },
  };
}

export function applyDefaultHeadersFooters(
  hf: ExcelHeaderFooter,
  headers?: Partial<{ left: string; center: string; right: string }>,
  footers?: Partial<{ left: string; center: string; right: string }>,
): void {
  if (headers) {
    if (headers.left !== undefined) {
      if (!("leftHeader" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.leftHeader is missing on host layout object",
        );
      }
      hf.leftHeader = headers.left;
    }
    if (headers.center !== undefined) {
      if (!("centerHeader" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.centerHeader is missing on host layout object",
        );
      }
      hf.centerHeader = headers.center;
    }
    if (headers.right !== undefined) {
      if (!("rightHeader" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.rightHeader is missing on host layout object",
        );
      }
      hf.rightHeader = headers.right;
    }
  }
  if (footers) {
    if (footers.left !== undefined) {
      if (!("leftFooter" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.leftFooter is missing on host layout object",
        );
      }
      hf.leftFooter = footers.left;
    }
    if (footers.center !== undefined) {
      if (!("centerFooter" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.centerFooter is missing on host layout object",
        );
      }
      hf.centerFooter = footers.center;
    }
    if (footers.right !== undefined) {
      if (!("rightFooter" in hf)) {
        throw new Error(
          "PageLayout.headersFooters.defaultForAllPages.rightFooter is missing on host layout object",
        );
      }
      hf.rightFooter = footers.right;
    }
  }
}

export function loadDefaultHeadersFooters(hf: ExcelHeaderFooter): void {
  hf.load(HEADER_FOOTER_LOAD);
}

export { HEADER_FOOTER_LOAD };
