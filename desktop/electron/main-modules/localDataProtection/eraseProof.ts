import { createHash } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";
import { app } from "electron";
import type { ManagedReplicaCategory } from "./managedReplicaRegistry";

export interface EraseProofReplicaResult {
  pathDigest: string;
  category: ManagedReplicaCategory;
  status: "erased" | "failed" | "skipped";
  error?: string;
}

export interface LocalDataEraseProof {
  formatVersion: 1;
  createdAt: string;
  installIdDigest: string;
  keyDestruction: {
    destroyedKeyIds: number[];
    keyMaterialDestroyed: boolean;
    error?: string;
  };
  replicas: EraseProofReplicaResult[];
}

function proofDirectory(userDataPath = app.getPath("userData")): string {
  return path.join(userDataPath, "erase-proofs");
}

export function writeLocalDataEraseProof(
  proof: LocalDataEraseProof,
  userDataPath?: string,
): { proofPath: string; proofDigest: string } {
  const dir = proofDirectory(userDataPath);
  fs.mkdirSync(dir, { recursive: true });
  const stamp = proof.createdAt.replace(/[:.]/g, "-");
  const proofPath = path.join(dir, `erase-proof-${stamp}.json`);
  const body = `${JSON.stringify(proof, null, 2)}\n`;
  fs.writeFileSync(proofPath, body, "utf8");
  const proofDigest = createHash("sha256").update(body, "utf8").digest("hex");
  return { proofPath, proofDigest };
}

export function buildEraseProofSummary(proof: LocalDataEraseProof, proofDigest: string) {
  return {
    createdAt: proof.createdAt,
    installIdDigest: proof.installIdDigest,
    proofDigest,
    destroyedKeyCount: proof.keyDestruction.keyMaterialDestroyed
      ? proof.keyDestruction.destroyedKeyIds.length
      : 0,
    keyMaterialDestroyed: proof.keyDestruction.keyMaterialDestroyed,
    replicaCount: proof.replicas.length,
    erasedCount: proof.replicas.filter((item) => item.status === "erased").length,
    failedCount: proof.replicas.filter((item) => item.status === "failed").length,
  };
}
