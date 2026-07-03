import { describe, expect, test } from "vitest";
import {
  buildImageDataUri,
  formatAttachmentSize,
  guessImageMimeType,
} from "./attachmentPreview";

describe("attachment preview helpers", () => {
  test("guesses image mime type from file extension", () => {
    expect(guessImageMimeType("C:\\Users\\29721\\Pictures\\invoice.webp")).toBe("image/webp");
    expect(guessImageMimeType("photo.jpeg")).toBe("image/jpeg");
    expect(guessImageMimeType("unknown")).toBe("image/png");
  });

  test("builds a browser-ready data URI", () => {
    expect(buildImageDataUri("AAAA", "image/png", "image.png")).toBe("data:image/png;base64,AAAA");
    expect(buildImageDataUri("BBBB", "application/octet-stream", "scan.jpg")).toBe("data:image/jpeg;base64,BBBB");
    expect(buildImageDataUri("data:image/gif;base64,CCCC", "image/gif")).toBe("data:image/gif;base64,CCCC");
  });

  test("formats attachment sizes compactly", () => {
    expect(formatAttachmentSize(512)).toBe("512 B");
    expect(formatAttachmentSize(1536)).toBe("1.5 KB");
    expect(formatAttachmentSize(120 * 1024)).toBe("120 KB");
    expect(formatAttachmentSize(2.5 * 1024 * 1024)).toBe("2.5 MB");
  });
});
