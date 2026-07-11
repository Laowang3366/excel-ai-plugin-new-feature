export function omitVersionMetadata(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value;
  const metadata = { ...(value as Record<string, unknown>) };
  delete metadata.version;
  return metadata;
}

export function toModelFacingSpreadsheetMetadata(value: unknown): unknown {
  const metadata = omitVersionMetadata(value);
  if (!metadata || typeof metadata !== "object" || Array.isArray(metadata)) return metadata;
  if ((metadata as Record<string, unknown>).host !== "wps") return metadata;

  return {
    ...metadata,
    formulaDialect: {
      regexFunction: "REGEXP",
      guidance: "WPS 正则提取使用 REGEXP；不要使用 Excel 方言的 REGEXEXTRACT/REGEXREPLACE/REGEXTEST 函数名",
    },
  };
}
