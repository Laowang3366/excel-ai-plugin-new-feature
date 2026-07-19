/** Parse A1-style cell (optional sheet!) into 0-based row/col. */
export function parseA1Cell(address: string): { row: number; col: number } | null {
  const bare = (address.includes("!") ? address.split("!")[1]! : address).replace(/\$/g, "");
  const match = /^([A-Za-z]+)(\d+)$/.exec(bare.trim());
  if (!match) return null;
  const letters = match[1]!.toUpperCase();
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { row: Number(match[2]) - 1, col: col - 1 };
}

export function toA1(row0: number, col0: number): string {
  let n = col0 + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return `${label}${row0 + 1}`;
}

/** Absolute A1 from a range origin (e.g. D5) + relative offsets. */
export function absoluteA1FromOrigin(originAddress: string, rowOffset: number, colOffset: number): string {
  const bare = (originAddress.includes("!") ? originAddress.split("!")[1]! : originAddress)
    .split(":")[0]!
    .replace(/\$/g, "");
  const origin = parseA1Cell(bare);
  if (!origin) return toA1(rowOffset, colOffset);
  return toA1(origin.row + rowOffset, origin.col + colOffset);
}

export function isSingleCellAddress(address: string): boolean {
  const bare = address.includes("!") ? address.split("!")[1]! : address;
  return !bare.includes(":") && !bare.includes(",");
}
