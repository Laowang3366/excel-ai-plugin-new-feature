import type { DataKeystore } from "./dataKeystore";
import {
  createRecordId,
  isProtectedBlob,
  openUtf8,
  parseProtectedRecordId,
  sealUtf8,
} from "./protectedBlob";

export interface PayloadProtection {
  currentKeyId(): number;
  protect(plaintext: string, aad: string, keyId?: number, recordId?: string): string;
  unprotect(value: string, aad: string): string;
  isProtected(value: string): boolean;
  recordIdOf(value: string): string | null;
}

class KeystorePayloadProtection implements PayloadProtection {
  constructor(private readonly keystore: DataKeystore) {}

  currentKeyId(): number {
    return this.keystore.currentKeyId;
  }

  protect(plaintext: string, aad: string, keyId?: number, recordId?: string): string {
    if (!plaintext) return plaintext;
    const target = keyId ?? this.keystore.currentKeyId;
    if (isProtectedBlob(plaintext)) {
      const existingRid = parseProtectedRecordId(plaintext) ?? createRecordId();
      const plain = this.unprotect(plaintext, aad);
      return sealUtf8(this.keystore.getKey(target), target, plain, aad, recordId ?? existingRid);
    }
    return sealUtf8(
      this.keystore.getKey(target),
      target,
      plaintext,
      aad,
      recordId ?? createRecordId(),
    );
  }

  unprotect(value: string, aad: string): string {
    if (!value || !isProtectedBlob(value)) return value;
    return openUtf8((id) => this.keystore.getKey(id), value, aad);
  }

  isProtected(value: string): boolean {
    return isProtectedBlob(value);
  }

  recordIdOf(value: string): string | null {
    return parseProtectedRecordId(value);
  }
}

let activeProtection: PayloadProtection | null = null;

export function setPayloadProtection(protection: PayloadProtection | null): void {
  activeProtection = protection;
}

export function getPayloadProtection(): PayloadProtection | null {
  return activeProtection;
}

export function createPayloadProtection(keystore: DataKeystore): PayloadProtection {
  return new KeystorePayloadProtection(keystore);
}
