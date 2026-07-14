import fs from "fs";
import path from "path";

export interface PathAuthorizerOptions {
  getDataPath: () => string;
  getPinnedFolders: () => string[];
  getExtraRoots?: () => string[];
}

export interface PathAuthorizer {
  authorizePath: (targetPath: string) => void;
  authorizeRoot: (targetPath: string) => void;
  isAuthorizedPath: (targetPath: string) => boolean;
}

function resolveCanonicalPath(targetPath: string): string {
  let current = path.resolve(targetPath);
  const missingSegments: string[] = [];

  while (!fs.existsSync(current)) {
    const parent = path.dirname(current);
    if (parent === current) break;
    missingSegments.unshift(path.basename(current));
    current = parent;
  }

  const canonicalParent = fs.existsSync(current) ? fs.realpathSync.native(current) : current;
  return path.resolve(canonicalParent, ...missingSegments).replace(/[\\/]+$/, "");
}

function normalizePathForCompare(targetPath: string): string {
  return resolveCanonicalPath(targetPath).toLowerCase();
}

function isPathInside(parentPath: string, childPath: string): boolean {
  const parent = normalizePathForCompare(parentPath);
  const child = normalizePathForCompare(childPath);
  return child === parent || child.startsWith(`${parent}${path.sep.toLowerCase()}`);
}

function compactPaths(paths: string[]): string[] {
  return paths
    .map((item) => typeof item === "string" ? item.trim() : "")
    .filter(Boolean);
}

export function createPathAuthorizer(options: PathAuthorizerOptions): PathAuthorizer {
  const authorizedPaths = new Set<string>();
  const authorizedRoots = new Set<string>();

  const getRuntimeRoots = () => compactPaths([
    options.getDataPath(),
    ...options.getPinnedFolders(),
    ...(options.getExtraRoots?.() ?? []),
  ]);

  return {
    authorizePath: (targetPath: string) => {
      if (targetPath.trim()) authorizedPaths.add(normalizePathForCompare(targetPath));
    },
    authorizeRoot: (targetPath: string) => {
      if (targetPath.trim()) authorizedRoots.add(normalizePathForCompare(targetPath));
    },
    isAuthorizedPath: (targetPath: string) => {
      if (!targetPath.trim()) return false;
      const normalized = normalizePathForCompare(targetPath);
      if (authorizedPaths.has(normalized)) return true;
      for (const root of authorizedRoots) {
        if (isPathInside(root, normalized)) return true;
      }
      return getRuntimeRoots().some((root) => isPathInside(root, normalized));
    },
  };
}

export function assertAuthorizedPath(authorizer: PathAuthorizer, targetPath: string): string {
  if (!authorizer.isAuthorizedPath(targetPath)) {
    throw new Error(`未授权访问路径：${targetPath}`);
  }
  return resolveCanonicalPath(targetPath);
}
