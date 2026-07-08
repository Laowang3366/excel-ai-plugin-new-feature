import test from "node:test";
import assert from "node:assert/strict";
import { parseJsonBody } from "../src/middleware/jsonBody.js";

test("parses UTF-8 JSON bodies without changing Chinese text", () => {
  const body = Buffer.from(JSON.stringify({ note: "正式授权卡密" }), "utf8");

  assert.deepEqual(parseJsonBody(body, "application/json"), { note: "正式授权卡密" });
});

test("falls back to GB18030 when JSON body bytes are not valid UTF-8", () => {
  const body = Buffer.from([
    0x7b, 0x22, 0x6e, 0x6f, 0x74, 0x65, 0x22, 0x3a, 0x22,
    0xd5, 0xfd, 0xca, 0xbd, 0xca, 0xda, 0xc8, 0xa8, 0xbf, 0xa8, 0xc3, 0xdc,
    0x22, 0x7d,
  ]);

  assert.deepEqual(parseJsonBody(body, "application/json"), { note: "正式授权卡密" });
});

test("honors explicit GBK-compatible charset on JSON requests", () => {
  const body = Buffer.from([
    0x7b, 0x22, 0x6e, 0x6f, 0x74, 0x65, 0x22, 0x3a, 0x22,
    0xd5, 0xfd, 0xca, 0xbd, 0xca, 0xda, 0xc8, 0xa8, 0xbf, 0xa8, 0xc3, 0xdc,
    0x22, 0x7d,
  ]);

  assert.deepEqual(parseJsonBody(body, "application/json; charset=gbk"), { note: "正式授权卡密" });
});
