import { describe, expect, test } from "vitest";
import {
  STREAMING_REASONING_BOTTOM_THRESHOLD,
  STREAMING_REASONING_LIVE_RENDER_LIMIT,
  getStreamingReasoningDistanceToBottom,
  getStreamingReasoningVisibleText,
  shouldFollowStreamingReasoning,
} from "./StreamingOutput";

describe("streaming reasoning scroll behavior", () => {
  test("treats the reasoning block as following only when near the bottom", () => {
    expect(shouldFollowStreamingReasoning(STREAMING_REASONING_BOTTOM_THRESHOLD)).toBe(true);
    expect(shouldFollowStreamingReasoning(STREAMING_REASONING_BOTTOM_THRESHOLD + 1)).toBe(false);
  });

  test("calculates distance to bottom for the inner reasoning scroller", () => {
    expect(getStreamingReasoningDistanceToBottom({
      scrollHeight: 600,
      scrollTop: 420,
      clientHeight: 150,
    })).toBe(30);
    expect(getStreamingReasoningDistanceToBottom({
      scrollHeight: 600,
      scrollTop: 520,
      clientHeight: 150,
    })).toBe(0);
  });

  test("limits live reasoning rendering to a tail preview while keeping a notice", () => {
    const reasoning = "A".repeat(STREAMING_REASONING_LIVE_RENDER_LIMIT + 12);
    const visible = getStreamingReasoningVisibleText(reasoning);

    expect(visible).toContain("已暂存前 12 字");
    expect(visible.endsWith("A".repeat(STREAMING_REASONING_LIVE_RENDER_LIMIT))).toBe(true);
    expect(visible.length).toBeLessThan(reasoning.length + 80);
  });
});
