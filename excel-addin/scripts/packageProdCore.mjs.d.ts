export function expandPackageVersion(version: string): string;
export function requireFourPartVersion(version: string): string;
export function deriveViteBaseFromBaseUrl(baseUrl: string): string;
export function assertViteBaseMatchesBaseUrl(
  viteBase: string,
  baseUrl: string,
): string;
export function resolvePackageInputs(input: {
  baseUrl: string;
  version?: string | null;
  viteBase?: string | null;
  packageJsonVersion: string;
}): {
  baseUrl: string;
  viteBase: string;
  version: string;
  packageJsonVersion: string;
};
export function collectLocalAssetRefs(html: string): string[];
export function assertIndexAssetsUnderBase(
  html: string,
  viteBase: string,
): string[];
export function isSensitiveRelativePath(relPath: string): boolean;
export function assertNoSensitiveDistPaths(relativePaths: string[]): void;
export function buildSha256Sums(
  files: Array<{ relativePath: string; content: Buffer | string }>,
): string;
export function buildBuildInfo(meta: {
  gitSha: string;
  packageVersion: string;
  manifestVersion: string;
  baseUrl: string;
  viteBase: string;
}): {
  gitSha: string;
  packageVersion: string;
  manifestVersion: string;
  baseUrl: string;
  viteBase: string;
};
export function makeArtifactName(version: string, gitSha: string): string;
export function parseCliArgs(argv: string[]): {
  baseUrl: string | null;
  version: string | null;
  viteBase: string | null;
  gitSha: string | null;
  skipBuild: boolean;
  distDir: string | null;
  rootDir: string | null;
  help?: boolean;
};
