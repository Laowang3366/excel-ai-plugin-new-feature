import * as fs from "fs";
import * as path from "path";
import type { IndexResult } from "./types";
import type { KnowledgeIndexer } from "./knowledgeIndexer";
import { createLogger } from "../../shared/logger";

const builtinKnowledgeLogger = createLogger("BuiltinKnowledge");

const BUILTIN_KNOWLEDGE_MANIFEST = "builtin-knowledge.json";

interface BuiltinKnowledgeManifest {
  files?: Array<{
    path?: string;
    sha256?: string;
  }>;
}

export async function indexBuiltinKnowledge(indexer: KnowledgeIndexer): Promise<IndexResult[]> {
  const root = resolveBuiltinKnowledgeRoot();
  if (!root) {
    builtinKnowledgeLogger.warn("No builtin knowledge root found");
    return [];
  }

  const files = readBuiltinKnowledgeManifest(root);
  const results: IndexResult[] = [];
  for (const file of files) {
    const filePath = path.join(root, file.path);
    if (!fs.existsSync(filePath)) {
      builtinKnowledgeLogger.warn("Builtin knowledge file missing", { filePath });
      continue;
    }

    const result = await indexer.indexFile(filePath, {
      skipUnchanged: true,
      knownFileHash: file.sha256,
    });
    results.push(result);

    if (!result.success) {
      builtinKnowledgeLogger.warn("Builtin knowledge indexing failed", {
        filePath,
        error: result.error,
      });
    }
  }

  return results;
}

function readBuiltinKnowledgeManifest(root: string): Array<{ path: string; sha256?: string }> {
  const manifestPath = path.join(root, BUILTIN_KNOWLEDGE_MANIFEST);
  try {
    const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf-8")) as BuiltinKnowledgeManifest;
    return (manifest.files || [])
      .filter((file): file is { path: string; sha256?: string } =>
        Boolean(file.path) && !file.path!.includes("..")
      );
  } catch (error) {
    builtinKnowledgeLogger.warn("Builtin knowledge manifest unavailable", {
      manifestPath,
      error: error instanceof Error ? error.message : String(error),
    });
    return [];
  }
}

function resolveBuiltinKnowledgeRoot(): string | null {
  const candidates = [
    path.join(process.cwd(), "public", "knowledge"),
    process.resourcesPath
      ? path.join(process.resourcesPath, "public", "knowledge")
      : "",
    path.join(__dirname, "..", "..", "public", "knowledge"),
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) return candidate;
  }
  return null;
}
