import type { RolloutItem, RolloutLine } from "../shared/types";
import { getPayloadProtection } from "../../main-modules/localDataProtection/payloadProtection";
import { createRecordId, jsonlLineAad } from "../../main-modules/localDataProtection/protectedBlob";

/** Seal rollout JSONL lines with stable recordId AAD, or leave plaintext if protection is off. */
export function sealRolloutJsonlLines(items: RolloutItem[], relativePath: string): string[] {
  const protection = getPayloadProtection();
  return items.map((item) => {
    const line: RolloutLine = {
      timestamp: new Date().toISOString(),
      item,
    };
    const plain = JSON.stringify(line);
    if (!protection) return `${plain}\n`;
    const rid = createRecordId();
    return `${protection.protect(plain, jsonlLineAad(relativePath, rid), undefined, rid)}\n`;
  });
}
