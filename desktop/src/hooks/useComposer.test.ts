import { afterEach, describe, expect, it, vi } from "vitest";

import {
  COMPOSER_INPUT_MAX_LENGTH,
  limitComposerInput,
  resolveDroppedFiles,
} from "./useComposer";

vi.mock("../services/ipcApi", () => {
  const mock = {
    file: {
      getPathForFile: vi.fn(),
      writeTempFile: vi.fn(),
    },
  };
  return {
    ipcApi: mock,
    createMockIpcApi: vi.fn((overrides = {}) => ({ ...mock, ...overrides })),
  };
});

const { ipcApi } = await import("../services/ipcApi");
const mockedFileApi = ipcApi.file as unknown as {
  getPathForFile: ReturnType<typeof vi.fn>;
  writeTempFile: ReturnType<typeof vi.fn>;
};

describe("limitComposerInput", () => {
  it("keeps input within the 50000 character limit", () => {
    const exact = "a".repeat(COMPOSER_INPUT_MAX_LENGTH);
    expect(limitComposerInput(exact)).toBe(exact);
    expect(limitComposerInput(`${exact}overflow`)).toHaveLength(COMPOSER_INPUT_MAX_LENGTH);
  });
});

function createFile(name: string, type: string, content = "demo"): File {
  return new File([content], name, { type });
}

describe("resolveDroppedFiles", () => {
  afterEach(() => {
    vi.clearAllMocks();
  });

  it("uses Electron webUtils file paths for dragged local files", async () => {
    const file = createFile("report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    mockedFileApi.getPathForFile.mockReturnValue("D:\\docs\\report.xlsx");

    const result = await resolveDroppedFiles([file]);

    expect(result).toEqual([
      {
        filePath: "D:\\docs\\report.xlsx",
        fileName: "report.xlsx",
        fileType: "document",
        size: file.size,
      },
    ]);
    expect(mockedFileApi.writeTempFile).not.toHaveBeenCalled();
  });

  it("writes a temporary file when dragged files have no local path", async () => {
    const file = createFile("scan.png", "image/png", "image-bytes");
    mockedFileApi.getPathForFile.mockReturnValue("");
    mockedFileApi.writeTempFile.mockResolvedValue({
      success: true,
      filePath: "C:\\Temp\\image-1.png",
    });

    const result = await resolveDroppedFiles([file]);

    expect(mockedFileApi.writeTempFile).toHaveBeenCalledWith({
      prefix: "image",
      suffix: ".png",
      data: expect.any(String),
    });
    expect(result).toEqual([
      {
        filePath: "C:\\Temp\\image-1.png",
        fileName: "scan.png",
        fileType: "image",
        size: file.size,
      },
    ]);
  });
});
