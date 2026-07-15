import { describe, expect, test, vi } from "vitest";
import { createDocumentDismissHandlers } from "./useDocumentDismiss";

describe("createDocumentDismissHandlers", () => {
  test("dismisses on outside pointer events", () => {
    const onDismiss = vi.fn();
    const handlers = createDocumentDismissHandlers({ onDismiss });

    handlers.handlePointerEvent({ target: { closest: () => null } } as unknown as MouseEvent);

    expect(onDismiss).toHaveBeenCalledOnce();
  });

  test("ignores pointer events inside matching selectors", () => {
    const onDismiss = vi.fn();
    const handlers = createDocumentDismissHandlers({
      onDismiss,
      ignoreSelectors: [".menu"],
    });

    handlers.handlePointerEvent({
      target: { closest: (selector: string) => (selector === ".menu" ? {} : null) },
    } as unknown as MouseEvent);

    expect(onDismiss).not.toHaveBeenCalled();
  });

  test("ignores pointer events inside boundary refs", () => {
    const onDismiss = vi.fn();
    const target = {};
    const boundary = {
      current: {
        contains: (node: unknown) => node === target,
      } as Element,
    };
    const handlers = createDocumentDismissHandlers({
      onDismiss,
      boundaryRefs: [boundary],
    });

    handlers.handlePointerEvent({ target } as unknown as MouseEvent);

    expect(onDismiss).not.toHaveBeenCalled();
  });

  test("dismisses on Escape only", () => {
    const onDismiss = vi.fn();
    const handlers = createDocumentDismissHandlers({ onDismiss });

    handlers.handleKeyDown({ key: "Enter" } as KeyboardEvent);
    handlers.handleKeyDown({ key: "Escape" } as KeyboardEvent);

    expect(onDismiss).toHaveBeenCalledOnce();
  });
});
