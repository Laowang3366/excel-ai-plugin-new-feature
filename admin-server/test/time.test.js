import test from "node:test";
import assert from "node:assert/strict";
import { formatBeijingDateTime, withBeijingDateTimes } from "../src/time.js";

test("formats SQLite UTC datetime strings as Beijing time", () => {
  assert.equal(formatBeijingDateTime("2026-07-07 12:56:14"), "2026-07-07 20:56:14");
});

test("keeps empty timestamp values unchanged", () => {
  assert.equal(formatBeijingDateTime(null), null);
  assert.equal(formatBeijingDateTime(""), "");
});

test("converts common response timestamp fields to Beijing time", () => {
  const record = withBeijingDateTimes({
    created_at: "2026-07-07 12:56:14",
    activated_at: "2026-07-07 12:56:14",
    last_heartbeat: "2026-07-07 12:56:14",
    note: "ok",
  });

  assert.deepEqual(record, {
    created_at: "2026-07-07 20:56:14",
    activated_at: "2026-07-07 20:56:14",
    last_heartbeat: "2026-07-07 20:56:14",
    note: "ok",
  });
});
