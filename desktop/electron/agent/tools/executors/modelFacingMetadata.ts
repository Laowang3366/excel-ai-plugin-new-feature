export function omitVersionMetadata(value: unknown): unknown {
  if (!value || typeof value !== "object" || Array.isArray(value)) return value;
  const metadata = { ...(value as Record<string, unknown>) };
  delete metadata.version;
  return metadata;
}
