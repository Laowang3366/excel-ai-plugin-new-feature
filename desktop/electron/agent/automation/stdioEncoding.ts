export function decodeProcessOutput(output: Buffer | string | null | undefined): string {
  if (output === null || output === undefined) return "";
  if (typeof output === "string") return output;

  const utf8 = new TextDecoder("utf-8", { fatal: false }).decode(output);
  try {
    const gb18030 = new TextDecoder("gb18030", { fatal: false }).decode(output);
    return shouldPreferLegacyDecode(utf8, gb18030) ? gb18030 : utf8;
  } catch {
    return utf8;
  }
}

function shouldPreferLegacyDecode(utf8: string, legacy: string): boolean {
  if (!legacy || legacy === utf8) return false;

  const utf8Score = mojibakeScore(utf8);
  const legacyScore = mojibakeScore(legacy);
  if (legacyScore < utf8Score) return true;

  return utf8Score === 0 &&
    legacyScore === 0 &&
    containsCjk(legacy) &&
    !containsCjk(utf8) &&
    legacyUtf8ArtifactScore(utf8) > 0;
}

function mojibakeScore(value: string): number {
  return countMatches(value, /\uFFFD/g) * 10 +
    countMatches(value, /锟|斤拷/g) * 8 +
    countMatches(value, /[鐢鍛宸瀹绋杈妯搴鏂枃璺緞涓]/g) +
    countMatches(value, /(?:â€|â€™|â€œ|â€�|â€“|â€”|Ã|Â)/g) * 2;
}

function legacyUtf8ArtifactScore(value: string): number {
  return ["ȡ", "Ŀ¼", "״̬", "ֹͣ"].reduce(
    (score, marker) => score + (value.includes(marker) ? 1 : 0),
    0
  );
}

function containsCjk(value: string): boolean {
  return /[\u3400-\u9fff]/.test(value);
}

function countMatches(value: string, pattern: RegExp): number {
  return value.match(pattern)?.length ?? 0;
}
