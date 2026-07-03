import { describe, expect, it } from "vitest";
import { decodeProcessOutput } from "./stdioEncoding";

describe("decodeProcessOutput", () => {
  it("keeps utf-8 output readable", () => {
    expect(decodeProcessOutput(Buffer.from("中文路径", "utf8"))).toBe("中文路径");
  });

  it("falls back to gb18030 for legacy Windows output", () => {
    const gbkChinese = Buffer.from([0xd6, 0xd0, 0xce, 0xc4]);
    expect(decodeProcessOutput(gbkChinese)).toBe("中文");
  });

  it("falls back when short GBK output is valid but suspicious utf-8", () => {
    const gbkDirectory = Buffer.from([0xc4, 0xbf, 0xc2, 0xbc]);
    const gbkStatus = Buffer.from([0xd7, 0xb4, 0xcc, 0xac]);

    expect(decodeProcessOutput(gbkDirectory)).toBe("目录");
    expect(decodeProcessOutput(gbkStatus)).toBe("状态");
  });
});
