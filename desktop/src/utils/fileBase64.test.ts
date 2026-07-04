import { afterEach, describe, expect, it, vi } from "vitest";
import { readFileAsBase64 } from "./fileBase64";

const originalFileReader = globalThis.FileReader;

describe("fileBase64", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    Object.defineProperty(globalThis, "FileReader", {
      configurable: true,
      writable: true,
      value: originalFileReader,
    });
  });

  it("reads a FileReader data URL and strips the prefix", async () => {
    class MockFileReader {
      result: string | ArrayBuffer | null = null;
      error: DOMException | null = null;
      onload: (() => void) | null = null;
      onerror: (() => void) | null = null;

      readAsDataURL(_file: File): void {
        this.result = "data:image/png;base64,aGVsbG8=";
        this.onload?.();
      }
    }

    Object.defineProperty(globalThis, "FileReader", {
      configurable: true,
      writable: true,
      value: MockFileReader,
    });

    await expect(readFileAsBase64({} as File)).resolves.toBe("aGVsbG8=");
  });

  it("falls back to arrayBuffer when FileReader is unavailable", async () => {
    Object.defineProperty(globalThis, "FileReader", {
      configurable: true,
      writable: true,
      value: undefined,
    });

    const file = {
      arrayBuffer: () => Promise.resolve(new Uint8Array([104, 101, 108, 108, 111]).buffer),
    } as File;

    await expect(readFileAsBase64(file)).resolves.toBe("aGVsbG8=");
  });
});
