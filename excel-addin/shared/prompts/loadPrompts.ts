import manifest from "./manifest.json";

export interface PromptManifestEntry {
  id: string;
  sourcePath: string;
  generatedPath: string;
  sha256: string;
  bytes: number;
  /**
   * desktop-identical: byte-synced from desktop templates (runtime reusable).
   * addin-adapted: composed via templates/ overlay; generated copy is reference only.
   */
  mode?: "desktop-identical" | "addin-adapted";
}

export interface PromptManifest {
  generatedAt: string;
  sourceRoot: string;
  note: string;
  files: PromptManifestEntry[];
}

export const PROMPT_MANIFEST = manifest as PromptManifest;

export const PROMPT_IDS = PROMPT_MANIFEST.files.map((file) => file.id);

const generatedModules = import.meta.glob("./generated/**/*.md", {
  query: "?raw",
  import: "default",
  eager: true,
}) as Record<string, string>;

const templateModules = import.meta.glob("./templates/**/*.md", {
  query: "?raw",
  import: "default",
  eager: true,
}) as Record<string, string>;

function toGeneratedKey(id: string): string {
  return `./generated/${id}`;
}

function toTemplateKey(id: string): string {
  return `./templates/${id}`;
}

export function listPromptIds(): string[] {
  return [...PROMPT_IDS];
}

/**
 * Prefer add-in adapted templates when present; otherwise desktop-synced generated text.
 * Adapted overlays keep Excel methodology while stripping Electron/COM/Word/PPT paths.
 */
export function getPromptText(id: string): string {
  const adapted = templateModules[toTemplateKey(id)];
  if (typeof adapted === "string") return adapted;
  const text = generatedModules[toGeneratedKey(id)];
  if (typeof text === "string") return text;
  throw new Error(`Prompt not found: ${id}`);
}

export function getPromptEntry(id: string): PromptManifestEntry {
  const entry = PROMPT_MANIFEST.files.find((file) => file.id === id);
  if (!entry) throw new Error(`Prompt manifest entry missing: ${id}`);
  return entry;
}

/** True when an add-in adapted template overrides the desktop-synced copy. */
export function hasAdaptedPrompt(id: string): boolean {
  return typeof templateModules[toTemplateKey(id)] === "string";
}

/** @deprecated Use promptComposer.renderPromptTemplate (throws on missing vars). */
export { renderPromptTemplate } from "./promptComposer";
