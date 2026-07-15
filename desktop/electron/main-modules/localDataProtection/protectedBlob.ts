import { randomBytes } from "node:crypto";
import { decryptAesGcm, encryptAesGcm } from "./aesGcm";

export const PROTECTED_BLOB_PREFIX = "ldp:v1:";
export const PROTECTED_BLOB_FORMAT_VERSION = 1;

export interface ProtectedBlobEnvelope {
  v: number;
  kid: number;
  /** Cleartext stable record identity; participates in AAD. */
  rid: string;
  iv: string;
  ct: string;
  tag: string;
}

export function createRecordId(): string {
  return randomBytes(12).toString("hex");
}

export function isProtectedBlob(value: string): boolean {
  return typeof value === "string" && value.startsWith(PROTECTED_BLOB_PREFIX);
}

function encodeEnvelope(envelope: ProtectedBlobEnvelope): string {
  return `${PROTECTED_BLOB_PREFIX}${Buffer.from(JSON.stringify(envelope), "utf8").toString("base64")}`;
}

function decodeEnvelope(sealed: string): ProtectedBlobEnvelope {
  const encoded = sealed.slice(PROTECTED_BLOB_PREFIX.length);
  try {
    return JSON.parse(Buffer.from(encoded, "base64").toString("utf8")) as ProtectedBlobEnvelope;
  } catch {
    throw new Error("invalid_protected_blob");
  }
}

export function parseProtectedEnvelope(sealed: string): ProtectedBlobEnvelope | null {
  if (!isProtectedBlob(sealed)) return null;
  try {
    const envelope = decodeEnvelope(sealed);
    if (envelope.v !== PROTECTED_BLOB_FORMAT_VERSION || !Number.isInteger(envelope.kid)) {
      return null;
    }
    return envelope;
  } catch {
    return null;
  }
}

export function parseProtectedKeyId(sealed: string): number | null {
  return parseProtectedEnvelope(sealed)?.kid ?? null;
}

export function parseProtectedRecordId(sealed: string): string | null {
  const rid = parseProtectedEnvelope(sealed)?.rid;
  return typeof rid === "string" && rid ? rid : null;
}

/** Seal with optional stable rid; rid is cleartext in envelope and must be in AAD. */
export function sealUtf8(
  key: Buffer,
  keyId: number,
  plaintext: string,
  aad: string,
  recordId = createRecordId(),
): string {
  const sealed = encryptAesGcm(
    key,
    Buffer.from(plaintext, "utf8"),
    aad ? Buffer.from(aad, "utf8") : undefined,
  );
  return encodeEnvelope({
    v: PROTECTED_BLOB_FORMAT_VERSION,
    kid: keyId,
    rid: recordId,
    iv: sealed.iv.toString("base64"),
    ct: sealed.ciphertext.toString("base64"),
    tag: sealed.tag.toString("base64"),
  });
}

export function openUtf8(
  resolveKey: (keyId: number) => Buffer,
  sealed: string,
  aad: string,
): string {
  if (!isProtectedBlob(sealed)) return sealed;
  const envelope = decodeEnvelope(sealed);
  if (envelope.v !== PROTECTED_BLOB_FORMAT_VERSION || !Number.isInteger(envelope.kid)) {
    throw new Error("unsupported_protected_blob_version");
  }
  if (!envelope.rid) throw new Error("missing_protected_record_id");
  const key = resolveKey(envelope.kid);
  const plaintext = decryptAesGcm(
    key,
    Buffer.from(envelope.iv, "base64"),
    Buffer.from(envelope.ct, "base64"),
    Buffer.from(envelope.tag, "base64"),
    aad ? Buffer.from(aad, "utf8") : undefined,
  );
  return plaintext.toString("utf8");
}

/** Build AAD for JSONL/archive lines: store.path.recordId.column */
export function jsonlLineAad(relativePath: string, recordId: string): string {
  return `sessions.jsonl.${relativePath}.${recordId}.line`;
}
