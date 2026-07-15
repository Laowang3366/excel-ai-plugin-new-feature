import * as fs from "node:fs";
import * as path from "node:path";

import {
  cleanupPreparedDataPathMigration,
  commitPreparedDataPathMigration,
  prepareDataPathMigration,
} from "./dataPathMigration";
import { isSensitiveFieldName } from "../shared/sensitiveData";

export interface UserDataExportResult {
  exportPath: string;
  exportedAt: string;
  categories: string[];
}

const EXPORT_CATEGORIES = [
  "settings-without-secrets",
  "sessions",
  "long-term-memory-and-goals",
  "knowledge-index",
  "office-backups-and-automation",
  "application-logs",
  "temporary-files",
];

export async function exportUserDataDirectory(options: {
  sourceDataPath: string;
  targetPath: string;
  sanitizedSettings: Record<string, unknown>;
  now?: Date;
}): Promise<UserDataExportResult> {
  const exportedAt = (options.now ?? new Date()).toISOString();
  const prepared = await prepareDataPathMigration(options.sourceDataPath, options.targetPath);
  let committed = false;
  try {
    const settingsDirectory = path.join(prepared.stageDataPath, "settings");
    await fs.promises.rm(settingsDirectory, { recursive: true, force: true });
    await fs.promises.mkdir(settingsDirectory, { recursive: true });
    await writeJson(path.join(settingsDirectory, "privacy-export-settings.json"), {
      exportedAt,
      credentials: "omitted",
      settings: redactExportSettings(options.sanitizedSettings),
    });
    await writeJson(path.join(prepared.stageDataPath, "privacy-export-manifest.json"), {
      schemaVersion: 1,
      exportedAt,
      categories: EXPORT_CATEGORIES,
      notes: [
        "Provider API keys, OCR tokens and custom header secrets are omitted or masked.",
        "Original files referenced by the knowledge index or conversations are not copied unless they were already stored inside the application data directory.",
      ],
    });
    await commitPreparedDataPathMigration(prepared);
    committed = true;
    return {
      exportPath: prepared.targetDataPath,
      exportedAt,
      categories: [...EXPORT_CATEGORIES],
    };
  } finally {
    await cleanupPreparedDataPathMigration(prepared, false);
    if (!committed && prepared.targetExisted) {
      await fs.promises.mkdir(prepared.targetDataPath, { recursive: true }).catch(() => {});
    }
  }
}

function redactExportSettings(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(redactExportSettings);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(
    Object.entries(value).map(([key, child]) => [
      key,
      isSensitiveFieldName(key) ? "[REDACTED]" : redactExportSettings(child),
    ]),
  );
}

async function writeJson(filePath: string, value: unknown): Promise<void> {
  await fs.promises.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, {
    encoding: "utf8",
    flag: "wx",
  });
}
