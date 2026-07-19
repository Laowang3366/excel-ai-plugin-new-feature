export const DEFAULT_APP_ID: string;
export const DEFAULT_VERSION: string;
export const DEFAULT_DEV_BASE_URL: string;
export const SUPPORT_URL: string;

export function escapeXmlAttr(value: string): string;
export function unescapeXmlAttr(value: string): string;
export function extractAttr(xml: string, name: string): string | null;
export function normalizeBaseUrl(input: string): string;
export function baseUrlOrigin(baseUrl: string): string;
export function joinBaseUrl(baseUrl: string, relPath: string): string;
export function isLocalhostHost(hostname: string): boolean;

export function renderOfficeManifest(opts: {
  mode: "dev" | "prod";
  baseUrl?: string;
  version?: string;
  appId?: string;
  template: string;
}): string;

export function validateOfficeManifest(
  xml: string,
  opts?: { mode?: "dev" | "prod" },
): { ok: boolean; errors: string[] };
