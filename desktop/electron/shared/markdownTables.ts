export function extractMarkdownTables(markdown: string): string[][] {
  const lines = markdown.split(/\r?\n/);
  const rows: string[][] = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!isMarkdownTableRow(line)) continue;
    const next = lines[i + 1]?.trim() || "";
    if (!isMarkdownTableSeparator(next)) continue;

    const tableRows: string[][] = [parseMarkdownTableRow(line)];
    i += 2;
    while (i < lines.length && isMarkdownTableRow(lines[i].trim())) {
      tableRows.push(parseMarkdownTableRow(lines[i].trim()));
      i += 1;
    }
    rows.push(...tableRows);
  }

  return rows;
}

function parseMarkdownTableRow(line: string): string[] {
  return line
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split("|")
    .map((cell) => cell.trim());
}

function isMarkdownTableRow(line: string): boolean {
  return line.startsWith("|") && line.endsWith("|") && line.includes("|");
}

function isMarkdownTableSeparator(line: string): boolean {
  return /^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?$/.test(line);
}
