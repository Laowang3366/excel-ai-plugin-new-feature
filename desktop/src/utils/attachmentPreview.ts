const IMAGE_MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  webp: "image/webp",
  bmp: "image/bmp",
  gif: "image/gif",
};

export function guessImageMimeType(filePathOrName: string, fallback = "image/png"): string {
  const ext = filePathOrName.split(/[\\/]/).pop()?.split(".").pop()?.toLowerCase() || "";
  return IMAGE_MIME_BY_EXT[ext] || fallback;
}

export function buildImageDataUri(data: string, mimeType?: string, filePathOrName = ""): string {
  if (data.startsWith("data:")) return data;
  const resolvedMime = mimeType?.startsWith("image/")
    ? mimeType
    : guessImageMimeType(filePathOrName);
  return `data:${resolvedMime};base64,${data}`;
}

export function formatAttachmentSize(size?: number): string {
  if (typeof size !== "number" || !Number.isFinite(size) || size < 0) return "";
  if (size < 1024) return `${size} B`;
  const kb = size / 1024;
  if (kb < 1024) return `${kb < 10 ? kb.toFixed(1) : Math.round(kb)} KB`;
  const mb = kb / 1024;
  return `${mb < 10 ? mb.toFixed(1) : Math.round(mb)} MB`;
}
