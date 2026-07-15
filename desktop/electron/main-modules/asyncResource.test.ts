import { describe, expect, it, vi } from "vitest";

import { AsyncResource } from "./asyncResource";

describe("AsyncResource", () => {
  it("deduplicates concurrent initialization and publishes one instance", async () => {
    let resolve!: (value: { id: number }) => void;
    const create = vi.fn(
      () =>
        new Promise<{ id: number }>((done) => {
          resolve = done;
        }),
    );
    const resource = new AsyncResource(create, vi.fn());

    const first = resource.get();
    const second = resource.get();
    expect(create).toHaveBeenCalledTimes(1);

    resolve({ id: 1 });
    await expect(Promise.all([first, second])).resolves.toEqual([{ id: 1 }, { id: 1 }]);
  });

  it("clears failed initialization so the next call can retry", async () => {
    const create = vi
      .fn()
      .mockRejectedValueOnce(new Error("init failed"))
      .mockResolvedValueOnce({ id: 2 });
    const resource = new AsyncResource<{ id: number }>(create, vi.fn());

    await expect(resource.get()).rejects.toThrow("init failed");
    await expect(resource.get()).resolves.toEqual({ id: 2 });
    expect(create).toHaveBeenCalledTimes(2);
  });

  it("waits for initialization before disposing and can be opened again", async () => {
    const dispose = vi.fn().mockResolvedValue(undefined);
    const create = vi.fn().mockResolvedValueOnce({ id: 1 }).mockResolvedValueOnce({ id: 2 });
    const resource = new AsyncResource(create, dispose);

    const instance = await resource.get();
    await resource.close();
    expect(dispose).toHaveBeenCalledWith(instance);
    await expect(resource.get()).resolves.toEqual({ id: 2 });
  });
});
