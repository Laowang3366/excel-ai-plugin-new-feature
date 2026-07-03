import { describe, expect, test } from "vitest";
import {
  AUTO_SCROLL_BOTTOM_THRESHOLD,
  JUMP_TO_LATEST_THRESHOLD,
  STREAM_AUTO_SCROLL_INTERVAL_MS,
  USER_SCROLL_AUTO_FOLLOW_PAUSE_MS,
  getDistanceToBottom,
  getVisibleMessageItems,
  isUserScrollPauseActive,
  shouldAutoFollowLatest,
  shouldRunScheduledAutoScroll,
  shouldShowJumpToLatest,
} from "./ChatMessageList";

describe("chat message list scroll behavior", () => {
  test("treats the message list as attached to latest content only near the bottom", () => {
    expect(shouldAutoFollowLatest(AUTO_SCROLL_BOTTOM_THRESHOLD)).toBe(true);
    expect(shouldAutoFollowLatest(AUTO_SCROLL_BOTTOM_THRESHOLD + 1)).toBe(false);
  });

  test("shows the jump button only after the user scrolls clearly away from latest content", () => {
    expect(shouldShowJumpToLatest(JUMP_TO_LATEST_THRESHOLD)).toBe(false);
    expect(shouldShowJumpToLatest(JUMP_TO_LATEST_THRESHOLD + 1)).toBe(true);
  });

  test("calculates distance to the bottom without returning negative values", () => {
    expect(getDistanceToBottom({ scrollHeight: 1000, scrollTop: 700, clientHeight: 200 })).toBe(100);
    expect(getDistanceToBottom({ scrollHeight: 1000, scrollTop: 900, clientHeight: 200 })).toBe(0);
  });

  test("runs scheduled auto-scroll only while streaming and still following latest content", () => {
    expect(STREAM_AUTO_SCROLL_INTERVAL_MS).toBeGreaterThanOrEqual(1000);
    expect(shouldRunScheduledAutoScroll({ isStreaming: true, shouldFollowLatest: true })).toBe(true);
    expect(shouldRunScheduledAutoScroll({
      isStreaming: true,
      shouldFollowLatest: true,
      userScrollPauseActive: true,
    })).toBe(false);
    expect(shouldRunScheduledAutoScroll({ isStreaming: true, shouldFollowLatest: false })).toBe(false);
    expect(shouldRunScheduledAutoScroll({ isStreaming: false, shouldFollowLatest: true })).toBe(false);
  });

  test("keeps auto-follow paused for a short window after the user scrolls away", () => {
    expect(isUserScrollPauseActive({
      now: 10_000,
      lastUserScrollAwayAt: 0,
    })).toBe(false);
    expect(isUserScrollPauseActive({
      now: 10_000,
      lastUserScrollAwayAt: 10_000 - USER_SCROLL_AUTO_FOLLOW_PAUSE_MS + 1,
    })).toBe(true);
    expect(isUserScrollPauseActive({
      now: 10_000,
      lastUserScrollAwayAt: 10_000 - USER_SCROLL_AUTO_FOLLOW_PAUSE_MS,
    })).toBe(false);
  });

  test("keeps only the latest message items in the render window", () => {
    const messages = Array.from({ length: 6 }, (_, index) => ({ id: `msg-${index}` }));

    expect(getVisibleMessageItems(messages, 3)).toEqual({
      visibleMessages: messages.slice(3),
      hiddenCount: 3,
    });
    expect(getVisibleMessageItems(messages, 6)).toEqual({
      visibleMessages: messages,
      hiddenCount: 0,
    });
  });
});
