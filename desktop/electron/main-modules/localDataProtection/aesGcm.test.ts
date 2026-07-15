import { describe, expect, it } from "vitest";
import { decryptAesGcm, encryptAesGcm, generateDataKey } from "./aesGcm";
import { isProtectedBlob, openUtf8, parseProtectedRecordId, sealUtf8 } from "./protectedBlob";

describe("local data AES-GCM envelopes", () => {
  it("round-trips utf8 payloads with recordId AAD and rejects tampering", () => {
    const key = generateDataKey();
    const rid = "rec-abc";
    const aad = `logs.rollout_events.${rid}.item_json`;
    const sealed = sealUtf8(key, 1, "secret-body", aad, rid);
    expect(isProtectedBlob(sealed)).toBe(true);
    expect(parseProtectedRecordId(sealed)).toBe(rid);
    expect(openUtf8(() => key, sealed, aad)).toBe("secret-body");
    expect(() => openUtf8(() => key, sealed, "logs.rollout_events.other.item_json")).toThrow();
    const corrupted = `${sealed.slice(0, -4)}abcd`;
    expect(() => openUtf8(() => key, corrupted, aad)).toThrow();
  });

  it("rejects ciphertext with wrong key", () => {
    const keyA = generateDataKey();
    const keyB = generateDataKey();
    const sealed = encryptAesGcm(keyA, Buffer.from("x"), Buffer.from("aad"));
    expect(() =>
      decryptAesGcm(keyB, sealed.iv, sealed.ciphertext, sealed.tag, Buffer.from("aad")),
    ).toThrow();
  });
});
