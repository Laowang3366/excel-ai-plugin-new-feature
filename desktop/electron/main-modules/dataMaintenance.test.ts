import { describe, expect, it } from "vitest";

import {
  assertDataMaintenanceAvailable,
  DATA_MAINTENANCE_BUSY_MESSAGE,
  guardDataOperation,
  hasActiveDataOperations,
} from "./dataMaintenance";

describe("data maintenance guard", () => {
  it("allows normal work and rejects work while the data root is under maintenance", () => {
    expect(() => assertDataMaintenanceAvailable(() => false)).not.toThrow();
    expect(() => assertDataMaintenanceAvailable(() => true)).toThrow(DATA_MAINTENANCE_BUSY_MESSAGE);
  });

  it("tracks an operation until its asynchronous work settles", async () => {
    let finish!: () => void;
    const pending = new Promise<void>((resolve) => {
      finish = resolve;
    });
    const operation = guardDataOperation(
      () => false,
      async () => pending,
    );

    const result = operation();
    expect(hasActiveDataOperations()).toBe(true);
    finish();
    await result;
    expect(hasActiveDataOperations()).toBe(false);
  });

  it("releases the active operation count after a failure", async () => {
    const operation = guardDataOperation(
      () => false,
      async () => {
        throw new Error("failed");
      },
    );

    await expect(operation()).rejects.toThrow("failed");
    expect(hasActiveDataOperations()).toBe(false);
  });
});
