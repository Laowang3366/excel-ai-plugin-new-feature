export interface FormatFileSizeOptions {
  emptyText?: string;
  compact?: boolean;
}

export function formatFileSize(size: number | undefined, options: FormatFileSizeOptions = {}): string {
  const { emptyText = "0 B", compact = false } = options;
  if (typeof size !== "number" || !Number.isFinite(size) || size < 0) return emptyText;
  if (size < 1024) return `${size} B`;

  const kb = size / 1024;
  if (kb < 1024) {
    return compact ? `${kb < 10 ? kb.toFixed(1) : Math.round(kb)} KB` : `${kb.toFixed(1)} KB`;
  }

  const mb = kb / 1024;
  return compact ? `${mb < 10 ? mb.toFixed(1) : Math.round(mb)} MB` : `${mb.toFixed(1)} MB`;
}
