export const RUNTIME_DESKTOP_DEP_PATTERNS: { id: string; re: RegExp }[];
export function findRuntimeDesktopDepHits(
  text: string,
  label?: string,
): { label: string; id: string }[];
export function isDocumentationOnlyDesktopMention(text: string): boolean;
export function collectRuntimeDesktopDepOffenders(
  files: { relativePath: string; content: string }[],
): string[];
export function assertNoRuntimeDesktopDepsInPackageFiles(
  files: { relativePath: string; content: string }[],
): void;
