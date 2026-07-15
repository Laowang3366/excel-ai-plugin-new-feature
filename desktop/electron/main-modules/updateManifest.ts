import { verify } from "node:crypto";

import { z } from "zod";

const SHA256_PATTERN = /^[a-f0-9]{64}$/iu;
const VERSION_PATTERN = /^\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/u;

export const InstallerUpdateSchema = z.object({
  url: z.string().url(),
  sha256: z.string().regex(SHA256_PATTERN),
  size: z.number().int().positive(),
});

export const HotPatchUpdateSchema = z.object({
  id: z.string().min(1).max(128).regex(/^[0-9A-Za-z._-]+$/u),
  baseVersion: z.string().regex(VERSION_PATTERN),
  sequence: z.number().int().positive(),
  publishedAt: z.string().datetime({ offset: true }),
  expiresAt: z.string().datetime({ offset: true }),
  url: z.string().url(),
  sha256: z.string().regex(SHA256_PATTERN),
  size: z.number().int().positive(),
  files: z.array(z.object({
    path: z.string().min(1).max(300),
    sha256: z.string().regex(SHA256_PATTERN),
    size: z.number().int().nonnegative(),
  }).strict()).min(1).max(2_000),
  restartRequired: z.literal(true),
}).strict();

export const HotPatchPolicySchema = z.object({
  revokedPatchIds: z.array(
    z.string().min(1).max(128).regex(/^[0-9A-Za-z._-]+$/u),
  ).max(2_000),
  minimumSafeSequenceByBaseVersion: z.record(
    z.string().regex(VERSION_PATTERN),
    z.number().int().nonnegative(),
  ),
}).strict();

export const RemoteUpdateManifestSchema = z.object({
  schemaVersion: z.literal(1),
  channel: z.literal("stable"),
  version: z.string().regex(VERSION_PATTERN),
  publishedAt: z.string().datetime({ offset: true }),
  releaseNotes: z.array(z.string().min(1).max(300)).max(30),
  installer: InstallerUpdateSchema.optional(),
  hotPatch: HotPatchUpdateSchema.optional(),
  hotPatchPolicy: HotPatchPolicySchema.optional(),
  signature: z.string().min(40),
}).strict();

export type InstallerUpdate = z.infer<typeof InstallerUpdateSchema>;
export type HotPatchUpdate = z.infer<typeof HotPatchUpdateSchema>;
export type HotPatchPolicy = z.infer<typeof HotPatchPolicySchema>;
export type RemoteUpdateManifest = z.infer<typeof RemoteUpdateManifestSchema>;

export function canonicalJson(value: unknown): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;

  const record = value as Record<string, unknown>;
  return `{${Object.keys(record)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${canonicalJson(record[key])}`)
    .join(",")}}`;
}

export function verifyRemoteUpdateManifest(
  input: unknown,
  publicKey: string | Buffer,
): RemoteUpdateManifest {
  const manifest = RemoteUpdateManifestSchema.parse(input);
  const { signature, ...unsignedManifest } = manifest;
  const valid = verify(
    null,
    Buffer.from(canonicalJson(unsignedManifest), "utf8"),
    publicKey,
    Buffer.from(signature, "base64"),
  );
  if (!valid) throw new Error("更新清单签名无效");
  return manifest;
}

export function compareVersions(left: string, right: string): number {
  const [leftCore, leftPre = ""] = left.split("-", 2);
  const [rightCore, rightPre = ""] = right.split("-", 2);
  const leftParts = leftCore.split(".").map(Number);
  const rightParts = rightCore.split(".").map(Number);

  for (let index = 0; index < Math.max(leftParts.length, rightParts.length); index += 1) {
    const difference = (leftParts[index] ?? 0) - (rightParts[index] ?? 0);
    if (difference !== 0) return difference > 0 ? 1 : -1;
  }
  if (leftPre === rightPre) return 0;
  if (!leftPre) return 1;
  if (!rightPre) return -1;
  return leftPre.localeCompare(rightPre, "en", { numeric: true });
}
