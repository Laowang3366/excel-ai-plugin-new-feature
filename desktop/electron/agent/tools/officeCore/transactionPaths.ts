import path from "node:path";

import type { OfficeActionInput, OfficeActionResult } from "./types";

export function listOfficeTransactionPaths(steps: OfficeActionInput[]): string[] {
  const paths = new Map<string, string>();
  for (const step of steps) {
    addPath(paths, step.filePath);
    addPath(paths, step.outputPath);
    const params = step.params || {};
    for (const key of ["outputPath", "wordOutputPath", "presentationOutputPath"] as const) {
      if (typeof params[key] === "string") addPath(paths, params[key]);
    }
    if (step.operation === "buildReportPackage") {
      const outputDirectory =
        typeof params.outputDirectory === "string" ? params.outputDirectory : step.outputPath;
      if (outputDirectory) {
        const baseName =
          typeof params.baseName === "string"
            ? params.baseName
            : `${path.basename(step.filePath || "report", path.extname(step.filePath || ""))}-报告`;
        addPath(paths, path.join(outputDirectory, `${baseName}.docx`));
        addPath(paths, path.join(outputDirectory, `${baseName}.pptx`));
      }
    }
  }
  return [...paths.values()];
}

export function collectResultArtifacts(result: OfficeActionResult): string[] {
  const artifacts = new Map<string, string>();
  if (result.outputPath && result.filePath && !samePath(result.outputPath, result.filePath))
    addPath(artifacts, result.outputPath);
  for (const change of result.changes) {
    if (change.target && path.isAbsolute(change.target) && path.extname(change.target))
      addPath(artifacts, change.target);
  }
  return [...artifacts.values()];
}

export function samePath(left: string, right: string): boolean {
  const a = path.resolve(left);
  const b = path.resolve(right);
  return process.platform === "win32" ? a.toLowerCase() === b.toLowerCase() : a === b;
}

function addPath(target: Map<string, string>, value?: string): void {
  if (!value || !path.isAbsolute(value)) return;
  const resolved = path.resolve(value);
  target.set(process.platform === "win32" ? resolved.toLowerCase() : resolved, resolved);
}
