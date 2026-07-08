import test from "node:test";
import assert from "node:assert/strict";
import { buildExportFilter, normalizeKeyIds } from "../src/keyAdmin.js";

test("normalizes key ids for bulk operations", () => {
  assert.deepEqual(normalizeKeyIds(["1", 2, "bad", 2, 0, -1, 3.4]), [1, 2, 3]);
});

test("builds unused key export filter", () => {
  assert.deepEqual(buildExportFilter("unused"), {
    label: "unused",
    where: "WHERE status = 'active' AND used_count = 0",
  });
});

test("builds active key export filter by default", () => {
  assert.deepEqual(buildExportFilter("active"), {
    label: "active",
    where: "WHERE status = 'active'",
  });
  assert.equal(buildExportFilter("unknown").label, "active");
});
