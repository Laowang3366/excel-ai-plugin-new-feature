// OpenAI-compatible providers often reject dots in function names.
export function sanitizeToolName(name: string): string {
  return name.replace(/\./g, "_");
}

export function desanitizeToolName(name: string): string {
  const compoundPrefixes: Record<string, string> = {
    office_action_: "office.action.",
  };
  for (const [safePrefix, internalPrefix] of Object.entries(compoundPrefixes)) {
    if (name.startsWith(safePrefix)) {
      return internalPrefix + name.slice(safePrefix.length);
    }
  }

  const prefixes = [
    "workbook", "range", "selection", "formula", "vba", "sheet",
    "ui", "file", "word", "presentation", "office", "knowledge",
    "web", "ocr", "memory",
  ];
  for (const prefix of prefixes) {
    if (name.startsWith(prefix + "_")) {
      return prefix + "." + name.slice(prefix.length + 1);
    }
  }
  return name;
}
