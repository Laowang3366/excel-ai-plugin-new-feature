import * as fs from "fs";
import * as path from "path";
import { promisify } from "util";
import type { Thread, ThreadId, ThreadMetadata } from "../shared/types";

const readdir = promisify(fs.readdir);

/** Get the default sessions root used by legacy local installs. */
export function getDefaultSessionsRoot(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || "C:\\", "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant", "sessions");
}

/** Build the dated JSONL rollout path for a newly created thread. */
export function getSessionFilePath(sessionsRoot: string, threadId: ThreadId): string {
  const now = new Date();
  const year = now.getFullYear().toString();
  const month = (now.getMonth() + 1).toString().padStart(2, "0");
  const day = now.getDate().toString().padStart(2, "0");

  const dir = path.join(sessionsRoot, year, month, day);
  const timestamp = now.toISOString().replace(/:/g, "-").replace(/\.\d+Z$/, "");
  const filename = `rollout-${timestamp}-${threadId}.jsonl`;

  return path.join(dir, filename);
}

export async function findAllRolloutFiles(sessionsRoot: string): Promise<string[]> {
  const files: string[] = [];
  await collectRolloutFiles(sessionsRoot, files);
  return files;
}

/** Find every live or compressed rollout artifact belonging to one thread. */
export async function findThreadRolloutArtifacts(
  sessionsRoot: string,
  threadId: ThreadId,
): Promise<string[]> {
  const files: string[] = [];
  const suffixes = [
    `-${threadId}.jsonl`,
    `-${threadId}.jsonl.gz`,
    `-${threadId}.jsonl.zst`,
  ];
  await collectMatchingRolloutArtifacts(sessionsRoot, suffixes, files);
  return files;
}

async function collectRolloutFiles(dir: string, files: string[]): Promise<void> {
  let entries: fs.Dirent[];
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectRolloutFiles(fullPath, files);
    } else if (entry.name.endsWith(".jsonl")) {
      files.push(fullPath);
    }
  }
}

async function collectMatchingRolloutArtifacts(
  dir: string,
  suffixes: string[],
  files: string[],
): Promise<void> {
  let entries: fs.Dirent[];
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await collectMatchingRolloutArtifacts(fullPath, suffixes, files);
    } else if (
      entry.isFile()
      && entry.name.startsWith("rollout-")
      && suffixes.some((suffix) => entry.name.endsWith(suffix))
    ) {
      files.push(fullPath);
    }
  }
}

export async function scanThreadMetadata(
  sessionsRoot: string,
  loadThreadByPath: (filePath: string) => Promise<Thread | null>
): Promise<ThreadMetadata[]> {
  const results: ThreadMetadata[] = [];
  await scanDirectory(sessionsRoot, results, loadThreadByPath);
  return results;
}

async function scanDirectory(
  dir: string,
  results: ThreadMetadata[],
  loadThreadByPath: (filePath: string) => Promise<Thread | null>
): Promise<void> {
  let entries: fs.Dirent[];
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      await scanDirectory(fullPath, results, loadThreadByPath);
    } else if (entry.name.endsWith(".jsonl") && entry.name.startsWith("rollout-")) {
      try {
        const thread = await loadThreadByPath(fullPath);
        if (thread) {
          results.push(thread.metadata);
        }
      } catch {
        // Skip damaged rollout files.
      }
    }
  }
}
