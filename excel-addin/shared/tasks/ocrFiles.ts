/** Browser-side OCR file acceptance (no desktop path IPC). */

const IMAGE_MIME = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/webp",
  "image/bmp",
  "image/gif",
]);

export function isImageOcrFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return (
    IMAGE_MIME.has(file.type) ||
    name.endsWith(".png") ||
    name.endsWith(".jpg") ||
    name.endsWith(".jpeg") ||
    name.endsWith(".webp") ||
    name.endsWith(".bmp") ||
    name.endsWith(".gif")
  );
}

export function isPdfOcrFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return file.type === "application/pdf" || name.endsWith(".pdf");
}

/** Accepted for upload UI; PDF may still be typed unsupported at recognize time. */
export function isAcceptedOcrFile(file: File): boolean {
  return isImageOcrFile(file) || isPdfOcrFile(file);
}

export function isLikelyInvoiceFile(file: File): boolean {
  return /发票|invoice|fapiao|票据/i.test(file.name);
}

export async function readFileAsBase64(file: File): Promise<string> {
  // Prefer arrayBuffer; fall back to FileReader for jsdom / partial File shims.
  if (typeof file.arrayBuffer === "function") {
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }
  return await new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(reader.error ?? new Error("FileReader failed"));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result === "string") {
        const comma = result.indexOf(",");
        resolve(comma >= 0 ? result.slice(comma + 1) : result);
        return;
      }
      if (result instanceof ArrayBuffer) {
        const bytes = new Uint8Array(result);
        let binary = "";
        for (let i = 0; i < bytes.length; i += 1) {
          binary += String.fromCharCode(bytes[i]!);
        }
        resolve(btoa(binary));
        return;
      }
      reject(new Error("unsupported FileReader result"));
    };
    reader.readAsDataURL(file);
  });
}

export function mimeTypeForFile(file: File): string {
  if (file.type) return file.type;
  const name = file.name.toLowerCase();
  if (name.endsWith(".png")) return "image/png";
  if (name.endsWith(".jpg") || name.endsWith(".jpeg")) return "image/jpeg";
  if (name.endsWith(".webp")) return "image/webp";
  if (name.endsWith(".bmp")) return "image/bmp";
  if (name.endsWith(".gif")) return "image/gif";
  if (name.endsWith(".pdf")) return "application/pdf";
  return "application/octet-stream";
}
