import manifest from "./manifest.json";

export interface PromptManifestEntry {
  id: string;
  sourcePath: string;
  generatedPath: string;
  sha256: string;
  bytes: number;
}

export interface PromptManifest {
  generatedAt: string;
  sourceRoot: string;
  note: string;
  files: PromptManifestEntry[];
}

export const PROMPT_MANIFEST = manifest as PromptManifest;

export const PROMPT_IDS = PROMPT_MANIFEST.files.map((file) => file.id);

const modules = import.meta.glob("./generated/**/*.md", {
  query: "?raw",
  import: "default",
  eager: true,
}) as Record<string, string>;

function toModuleKey(id: string): string {
  return `./generated/${id}`;
}

export function listPromptIds(): string[] {
  return [...PROMPT_IDS];
}

export function getPromptText(id: string): string {
  const key = toModuleKey(id);
  const text = modules[key];
  if (typeof text !== "string") {
    throw new Error(`Prompt not found: ${id}`);
  }
  return text;
}

export function getPromptEntry(id: string): PromptManifestEntry {
  const entry = PROMPT_MANIFEST.files.find((file) => file.id === id);
  if (!entry) throw new Error(`Prompt manifest entry missing: ${id}`);
  return entry;
}

/** @deprecated Use promptComposer.renderPromptTemplate (throws on missing vars). */
export { renderPromptTemplate } from "./promptComposer";
