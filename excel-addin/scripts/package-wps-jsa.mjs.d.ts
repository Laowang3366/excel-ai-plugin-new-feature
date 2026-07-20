export type WpsPackageArgs = {
  gitSha?: string | null;
  skipBuild?: boolean;
  distDir?: string | null;
  rootDir?: string | null;
};

export type WpsPackageSummary = {
  ok: true;
  artifactName: string;
  version: string;
  gitSha: string;
  distDir: string;
  addonDirectory: string;
  files: string[];
};

export function createWpsPackage(
  args?: WpsPackageArgs,
  env?: NodeJS.ProcessEnv,
): WpsPackageSummary;
