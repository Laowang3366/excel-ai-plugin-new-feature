export const WPS_ADDON_NAME: string;
export const WPS_ADDON_DIRECTORY: string;
export const WPS_ENTRY_SCRIPT: string;
export const WPS_PUBLISH_URL: string;

export type Validation = { ok: boolean; errors: string[] };
export type IndexValidation = Validation & { assets: string[] };

export function renderWpsPublishXml(): string;
export function validateWpsManifest(xml: string): Validation;
export function validateWpsRibbon(xml: string): Validation;
export function validateWpsEntryScript(source: string): Validation;
export function validateWpsPublishXml(xml: string): Validation;
export function prepareWpsIndexHtml(html: string): string;
export function validateWpsIndexHtml(html: string): IndexValidation;
export function normalizeWpsGitSha(value: string | null | undefined): string;
export function makeWpsArtifactName(version: string, gitSha: string): string;
export function validateWpsSourceBundle(bundle: {
  manifestXml: string;
  ribbonXml: string;
  entryScript: string;
  publishXml: string;
}): Validation;
