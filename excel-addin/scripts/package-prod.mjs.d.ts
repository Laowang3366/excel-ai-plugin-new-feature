export type PackageCliArgs = {
  baseUrl: string | null;
  version?: string | null;
  viteBase?: string | null;
  gitSha?: string | null;
  skipBuild?: boolean;
  distDir?: string | null;
  rootDir?: string | null;
};

export type PackageSummary = {
  ok: true;
  artifactName: string;
  baseUrl: string;
  viteBase: string;
  version: string;
  gitSha: string;
  distDir: string;
  files: string[];
};

export function createPackage(
  args: PackageCliArgs,
  env?: NodeJS.ProcessEnv,
): PackageSummary;
