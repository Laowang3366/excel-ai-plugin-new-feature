import test from "node:test";
import assert from "node:assert/strict";
import {
  assertBoundMachine,
  normalizeDevicePayload,
} from "../src/selfServiceDevices.js";

test("normalizes device self-service payload", () => {
  assert.deepEqual(normalizeDevicePayload({
    key: " xqk2-hcjj-t8ys-4lzw ",
    machine_id: " pc-001 ",
    target_machine_id: " pc-002 ",
  }), {
    key: "XQK2-HCJJ-T8YS-4LZW",
    machineId: "pc-001",
    targetMachineId: "pc-002",
  });
});

test("requires requester machine to already be bound to the key", () => {
  const binding = assertBoundMachine({
    key_code: "KEY-001",
    status: "active",
    id: 7,
    machine_id: "pc-001",
  });

  assert.equal(binding.keyId, 7);
});

test("rejects unbound requester machine", () => {
  assert.throws(
    () => assertBoundMachine(null),
    /当前设备未绑定该卡密/
  );
});

test("rejects disabled keys for device self-service", () => {
  assert.throws(
    () => assertBoundMachine({
      key_code: "KEY-001",
      status: "disabled",
      id: 7,
      machine_id: "pc-001",
    }),
    /卡密已被禁用/
  );
});
