export function flattenJson(value: unknown, pathKey = "$", lines: string[] = []): string[] {
  if (Array.isArray(value)) {
    if (value.length === 0) {
      lines.push(`${pathKey}: []`);
      return lines;
    }
    value.forEach((item, index) => flattenJson(item, `${pathKey}[${index}]`, lines));
    return lines;
  }

  if (value && typeof value === "object") {
    const entries = Object.entries(value as Record<string, unknown>);
    if (entries.length === 0) {
      lines.push(`${pathKey}: {}`);
      return lines;
    }
    for (const [key, child] of entries) {
      const childPath = /^[A-Za-z_$][\w$]*$/.test(key)
        ? `${pathKey}.${key}`
        : `${pathKey}[${JSON.stringify(key)}]`;
      flattenJson(child, childPath, lines);
    }
    return lines;
  }

  lines.push(`${pathKey}: ${jsonScalarToText(value)}`);
  return lines;
}

function jsonScalarToText(value: unknown): string {
  if (typeof value === "string") return value;
  if (value === null) return "null";
  return String(value);
}
