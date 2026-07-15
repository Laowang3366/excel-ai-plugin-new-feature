import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";

export const AES_GCM_ALGORITHM = "aes-256-gcm";
export const AES_GCM_KEY_BYTES = 32;
export const AES_GCM_IV_BYTES = 12;
export const AES_GCM_TAG_BYTES = 16;

export function generateDataKey(): Buffer {
  return randomBytes(AES_GCM_KEY_BYTES);
}

export function encryptAesGcm(
  key: Buffer,
  plaintext: Buffer,
  aad?: Buffer,
): { iv: Buffer; ciphertext: Buffer; tag: Buffer } {
  if (key.length !== AES_GCM_KEY_BYTES) {
    throw new Error("invalid_data_key_length");
  }
  const iv = randomBytes(AES_GCM_IV_BYTES);
  const cipher = createCipheriv(AES_GCM_ALGORITHM, key, iv);
  if (aad && aad.length > 0) cipher.setAAD(aad);
  const ciphertext = Buffer.concat([cipher.update(plaintext), cipher.final()]);
  return { iv, ciphertext, tag: cipher.getAuthTag() };
}

export function decryptAesGcm(
  key: Buffer,
  iv: Buffer,
  ciphertext: Buffer,
  tag: Buffer,
  aad?: Buffer,
): Buffer {
  if (key.length !== AES_GCM_KEY_BYTES) {
    throw new Error("invalid_data_key_length");
  }
  if (iv.length !== AES_GCM_IV_BYTES) {
    throw new Error("invalid_iv_length");
  }
  if (tag.length !== AES_GCM_TAG_BYTES) {
    throw new Error("invalid_tag_length");
  }
  const decipher = createDecipheriv(AES_GCM_ALGORITHM, key, iv);
  if (aad && aad.length > 0) decipher.setAAD(aad);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ciphertext), decipher.final()]);
}
