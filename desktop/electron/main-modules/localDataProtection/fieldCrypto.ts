import { getPayloadProtection } from "./payloadProtection";
import { isProtectedBlob } from "./protectedBlob";

/** AAD = store.table.rowId.column — binds ciphertext to record identity. */
export function fieldAad(store: string, table: string, rowId: string, column: string): string {
  return `${store}.${table}.${rowId}.${column}`;
}

export function protectFieldValue(
  value: string | null | undefined,
  aad: string,
  options?: { keyId?: number; recordId?: string },
): string | null {
  if (value == null) return null;
  if (value === "") return value;
  const protection = getPayloadProtection();
  if (!protection) return value;
  return protection.protect(value, aad, options?.keyId, options?.recordId);
}

export function unprotectFieldValue(value: string | null | undefined, aad: string): string | null {
  if (value == null) return null;
  if (value === "") return value;
  const protection = getPayloadProtection();
  if (!protection) return value;
  if (isProtectedBlob(value)) {
    return protection.unprotect(value, aad);
  }
  return value;
}

export function protectRequiredField(
  value: string,
  aad: string,
  options?: { keyId?: number; recordId?: string },
): string {
  return protectFieldValue(value, aad, options) ?? "";
}

export function unprotectRequiredField(value: string, aad: string): string {
  return unprotectFieldValue(value, aad) ?? "";
}
