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
export function assertIndexAssetsUnderBase(
  html: string,
  viteBase: string,
): string[];
export function assertLocalAssetFiles(
  distDir: string,
  relativePaths: string[],
): string[];
export function listFilesRecursiveStrict(dir: string, base?: string): string[];
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
export function formatSpawnFailure(result: {
  error?: { code?: string } | null;
  signal?: string | null;
  status?: number | null;
}): string;
export function parseCliArgs(argv: string[]): {
  baseUrl: string | null;
  version: string | null;
  viteBase: string | null;
  gitSha: string | null;
  help?: boolean;
};

export const OFFICE_JS_CDN_URL: string;
export function assertProductionDistClean(opts: {
  distDir: string;
  baseUrl: string;
  viteBase: string;
  relativePaths?: string[];
}): void;
