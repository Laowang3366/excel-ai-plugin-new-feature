import { afterEach, describe, expect, it, vi } from "vitest";
import type { AgentTurnInput, Turn } from "../../shared/types";
import { completeTurn, createTurn, createUserMessageItem } from "./turnRunner";

describe("turnRunner", () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("creates an in-progress turn for the target thread", () => {
    vi.spyOn(Date, "now").mockReturnValue(123456);

    const turn = createTurn("thread-1");

    expect(turn).toMatchObject({
      threadId: "thread-1",
      status: "in_progress",
      items: [],
      startedAt: 123456,
    });
    expect(turn.turnId).toMatch(/^turn-123456-/);
    expect(turn.completedAt).toBeUndefined();
  });

  it("creates a user message item from turn input without dropping attachments or client id", () => {
    vi.spyOn(Date, "now").mockReturnValue(234567);
    const input: AgentTurnInput = {
      content: "analyze this workbook",
      clientId: "client-message-1",
      attachments: [
        {
          filePath: "C:\\docs\\sample.xlsx",
          fileName: "sample.xlsx",
          fileType: "document",
          size: 1024,
        },
      ],
    };

    const item = createUserMessageItem(input);

    expect(item).toEqual({
      type: "user_message",
      id: "msg-234567",
      content: "analyze this workbook",
      attachments: input.attachments,
      clientId: "client-message-1",
      timestamp: 234567,
    });
  });

  it("marks a turn completed and records completion time", () => {
    const turn: Turn = {
      turnId: "turn-1",
      threadId: "thread-1",
      status: "in_progress",
      items: [],
      startedAt: 1000,
    };
    vi.spyOn(Date, "now").mockReturnValue(2000);

    completeTurn(turn);

    expect(turn.status).toBe("completed");
    expect(turn.completedAt).toBe(2000);
  });
});
