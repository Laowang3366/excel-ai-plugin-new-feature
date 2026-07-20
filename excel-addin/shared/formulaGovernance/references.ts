/**
 * Desktop-aligned A1 reference extraction (ExcelFormulaActionService regex).
 * String literals are blanked before parsing.
 */

import { makeCellId, normalizeA1Address, unescapeSheetName } from "./address";
import type { FormulaEdge, FormulaEdgeKind } from "./types";

const STRING_LITERAL_RE = /"(?:[^"]|"")*"/g;

/** [Book.xlsx]Sheet!A1 or [Book.xlsx]'Sheet'!A1 — desktop ExternalReferenceRegex */
const EXTERNAL_RE =
  /\[([^\]]+)\](?:'((?:[^']|'')+)'|([^'!][^!]*))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/gi;

/** Sheet1!A1 or 'My Sheet'!$B$2:$C$3 — desktop QualifiedReferenceRegex */
const QUALIFIED_RE =
  /(?<![A-Za-z0-9_.])(?:'((?:[^']|'')+)'|([A-Za-z_\u4e00-\u9fff][A-Za-z0-9_ .\u4e00-\u9fff-]*))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/gi;

/** Local A1 / A1:B2 — desktop LocalReferenceRegex */
const LOCAL_RE =
  /(?<![A-Za-z0-9_.!])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/gi;

export function removeStringLiterals(formula: string): string {
  return formula.replace(STRING_LITERAL_RE, '""');
}

export interface ParsedReference {
  kind: FormulaEdgeKind;
  targetId: string;
  reference: string;
  sheet?: string;
  address: string;
  workbook?: string;
}

function pushUnique(
  out: ParsedReference[],
  seen: Set<string>,
  item: ParsedReference,
): void {
  // Dedupe by kind+target (desktop edges DistinctBy From/To/Kind; A1 vs $A$1 same target)
  const key = `${item.kind}|${item.targetId}`;
  if (seen.has(key)) return;
  seen.add(key);
  out.push(item);
}

/**
 * Parse formula references from a cell on `sheetName`.
 * Order: external → cross-sheet → same-sheet (after stripping qualified).
 */
export function parseFormulaReferences(
  formula: string,
  sheetName: string,
): ParsedReference[] {
  const out: ParsedReference[] = [];
  const seen = new Set<string>();
  const stripped = removeStringLiterals(formula);

  for (const match of stripped.matchAll(EXTERNAL_RE)) {
    const workbook = match[1] ?? "";
    const sheet = unescapeSheetName((match[2] || match[3] || "").trim());
    const address = normalizeA1Address(match[4] ?? "");
    pushUnique(out, seen, {
      kind: "external",
      targetId: `external:[${workbook}]${sheet}!${address}`,
      reference: match[0] ?? "",
      workbook,
      sheet,
      address,
    });
  }

  for (const match of stripped.matchAll(QUALIFIED_RE)) {
    const idx = match.index ?? 0;
    // Skip the sheet!addr portion of an already-matched [Book]Sheet!addr
    if (idx > 0 && stripped[idx - 1] === "]") continue;
    const sheet = unescapeSheetName((match[1] || match[2] || "").trim());
    const address = normalizeA1Address(match[3] ?? "");
    if (!sheet || !address) continue;
    pushUnique(out, seen, {
      kind: "cross-sheet",
      targetId: makeCellId(sheet, address),
      reference: match[0] ?? "",
      sheet,
      address,
    });
  }

  const localHaystack = stripped.replace(QUALIFIED_RE, " ");
  for (const match of localHaystack.matchAll(LOCAL_RE)) {
    const address = normalizeA1Address(match[1] ?? "");
    if (!address) continue;
    pushUnique(out, seen, {
      kind: "same-sheet",
      targetId: makeCellId(sheetName, address),
      reference: (match[0] ?? address).trim(),
      sheet: sheetName,
      address,
    });
  }

  return out;
}

export function referencesToEdges(
  fromId: string,
  refs: ParsedReference[],
): FormulaEdge[] {
  return refs.map((ref) => ({
    from: fromId,
    to: ref.targetId,
    kind: ref.kind,
    reference: ref.reference,
  }));
}
