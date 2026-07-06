import { afterEach, describe, expect, it, vi } from "vitest";

import { getAttachmentFileType, resolveDroppedFiles } from "./composerAttachmentFiles";

vi.mock("../services/ipcApi", () => {
  const mock = {
    file: {
      getPathForFile: vi.fn(),
      writeTempFile: vi.fn(),
    },
  };
  return {
    ipcApi: mock,
  };
});

const { ipcApi } = await import("../services/ipcApi");
const mockedFileApi = ipcApi.file as unknown as {
  getPathForFile: ReturnType<typeof vi.fn>;
  writeTempFile: ReturnType<typeof vi.fn>;
};

function createFile(name: string, type: string, content = "demo"): File {
  return new File([content], name, { type });
}

describe("composerAttachmentFiles", () => {
  afterEach(() => {
    vi.clearAllMocks();
  });

  it("detects image attachments from extension when MIME is missing", () => {
    expect(getAttachmentFileType(createFile("paste.webp", ""))).toBe("image");
    expect(getAttachmentFileType(createFile("notes.docx", ""))).toBe("document");
  });

  it("uses local file paths before writing temporary files", async () => {
    const file = createFile("report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    mockedFileApi.getPathForFile.mockReturnValue("D:\\docs\\report.xlsx");

    await expect(resolveDroppedFiles([file])).resolves.toEqual([
      {
        filePath: "D:\\docs\\report.xlsx",
        fileName: "report.xlsx",
        fileType: "document",
        size: file.size,
      },
    ]);
    expect(mockedFileApi.writeTempFile).not.toHaveBeenCalled();
  });

  it("writes pathless image files to the temp file bridge", async () => {
    const file = createFile("scan.png", "image/png", "image-bytes");
    mockedFileApi.getPathForFile.mockReturnValue("");
    mockedFileApi.writeTempFile.mockResolvedValue({
      success: true,
      filePath: "C:\\Temp\\image-1.png",
    });

    await expect(resolveDroppedFiles([file])).resolves.toEqual([
      {
        filePath: "C:\\Temp\\image-1.png",
        fileName: "scan.png",
        fileType: "image",
        size: file.size,
      },
    ]);
    expect(mockedFileApi.writeTempFile).toHaveBeenCalledWith({
      prefix: "image",
      suffix: ".png",
      data: expect.any(String),
    });
  });
});
