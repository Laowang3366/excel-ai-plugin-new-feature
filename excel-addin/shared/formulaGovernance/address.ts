/** Cell id and A1 address helpers for formula governance. */

/** Strip sheet prefix and absolute $ markers for stable ids. */
export function normalizeA1Address(address: string): string {
  const bare = address.includes("!")
    ? address.slice(address.lastIndexOf("!") + 1)
    : address;
  return bare.replace(/\$/g, "").trim().toUpperCase();
}

export function makeCellId(sheetName: string, address: string): string {
  const sheet = sheetName.trim();
  const addr = normalizeA1Address(address);
  return `${sheet}!${addr}`;
}

export function parseCellId(
  id: string,
): { sheetName: string; address: string } | null {
  const bang = id.lastIndexOf("!");
  if (bang <= 0 || bang === id.length - 1) return null;
  return {
    sheetName: id.slice(0, bang),
    address: normalizeA1Address(id.slice(bang + 1)),
  };
}

/** True if text looks like an A1 cell or A1:A1 range (optional $). */
export function isA1Like(address: string): boolean {
  return /^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(address.trim());
}

export function unescapeSheetName(name: string): string {
  return name.replace(/''/g, "'");
}
