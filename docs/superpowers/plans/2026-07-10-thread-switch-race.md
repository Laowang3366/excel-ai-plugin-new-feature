# Thread Switch Race Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Prevent an older conversation navigation response or error from overwriting the user's newest thread switch, new-thread request, or active-thread deletion.

**Architecture:** Keep IPC calls and patch construction in `threadActions.ts`. Add a monotonically increasing navigation sequence in the Zustand store orchestration layer, where concurrent thread switches, new-thread requests, and active-thread deletions can be ordered. Only the latest navigation request may apply patches or errors. Track the pending switch target so deleting that target invalidates its load without cancelling switches when an unrelated thread is deleted. Re-check active and pending targets when deletion finishes because a thread can become selected after its delete request starts.

**Tech Stack:** TypeScript, Zustand, Vitest, Electron renderer IPC wrapper.

---

### Task 1: Reproduce Out-of-Order Conversation Loads

**Files:**
- Modify: `desktop/src/store/chatStore.test.ts`

- [x] **Step 1: Add a controllable thread load mock**

Extend the hoisted mock and mocked API:

```ts
const ipcMocks = vi.hoisted(() => ({
  // existing mocks
  loadThread: vi.fn(),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    // existing APIs
    thread: {
      resume: ipcMocks.resumeThread,
      list: ipcMocks.listThreads,
      load: ipcMocks.loadThread,
      newThread: ipcMocks.newThread,
    },
  },
}));
```

Add a deferred Promise helper:

```ts
function deferred<T>() {
  let resolve!: (value: T) => void;
  let reject!: (reason?: unknown) => void;
  const promise = new Promise<T>((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}
```

- [x] **Step 2: Write the failing stale-success test**

```ts
describe("chatStore switchThread", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useChatStore.setState({
      activeThreadId: null,
      messages: [],
      isStreaming: false,
      turnStatus: "idle",
      error: null,
    });
  });

  it("keeps the newest thread when an older load finishes last", async () => {
    const loadA = deferred<any>();
    const loadB = deferred<any>();
    ipcMocks.loadThread.mockImplementation((threadId: string) =>
      threadId === "thread-a" ? loadA.promise : loadB.promise
    );

    const switchA = useChatStore.getState().switchThread("thread-a");
    const switchB = useChatStore.getState().switchThread("thread-b");

    loadB.resolve({
      items: [{
        type: "assistant_message",
        id: "message-b",
        content: "会话 B",
        timestamp: 2,
      }],
    });
    await switchB;

    loadA.resolve({
      items: [{
        type: "assistant_message",
        id: "message-a",
        content: "会话 A",
        timestamp: 1,
      }],
    });
    await switchA;

    expect(useChatStore.getState().activeThreadId).toBe("thread-b");
    expect(useChatStore.getState().messages).toMatchObject([
      { id: "message-b", content: "会话 B" },
    ]);
  });
});
```

- [x] **Step 3: Write the failing stale-error test**

```ts
it("ignores an older load error after a newer switch succeeds", async () => {
  const loadA = deferred<any>();
  const loadB = deferred<any>();
  ipcMocks.loadThread.mockImplementation((threadId: string) =>
    threadId === "thread-a" ? loadA.promise : loadB.promise
  );

  const switchA = useChatStore.getState().switchThread("thread-a");
  const switchB = useChatStore.getState().switchThread("thread-b");

  loadB.resolve({ items: [] });
  await switchB;
  loadA.reject(new Error("A 加载失败"));
  await switchA;

  expect(useChatStore.getState().activeThreadId).toBe("thread-b");
  expect(useChatStore.getState().error).toBeNull();
});
```

- [x] **Step 4: Run the tests and verify RED**

Run:

```powershell
cd D:\excel-ai-plugin-new-feature\desktop
npx vitest run src/store/chatStore.test.ts
```

Expected: both new tests fail because the older A result or error is still applied after B.

### Task 2: Gate State Writes by the Latest Switch Request

**Files:**
- Modify: `desktop/src/store/chatStore.ts`
- Test: `desktop/src/store/chatStore.test.ts`

- [x] **Step 1: Add the navigation sequence**

Place the sequence beside the store helper functions:

```ts
let latestThreadNavigationRequestId = 0;
let pendingThreadSwitchTargetId: string | null = null;
```

- [x] **Step 2: Ignore stale results and errors**

Start each navigation with a new request id, and apply its result only if it is still latest:

```ts
function beginThreadNavigation(targetThreadId: string | null): number {
  pendingThreadSwitchTargetId = targetThreadId;
  return ++latestThreadNavigationRequestId;
}

function completeThreadNavigation(requestId: number): boolean {
  if (requestId !== latestThreadNavigationRequestId) {
    return false;
  }
  pendingThreadSwitchTargetId = null;
  return true;
}
```

Use the same gate for `switchThread` and `createNewThread`. For `deleteThread`, use it only when deleting the current thread or the pending switch target, so deleting an unrelated thread does not cancel the switch.

- [x] **Step 3: Cover cross-action navigation races**

Add regression tests for:

- A stale load after creating a new thread.
- A stale load after deleting the current thread.
- Deleting the pending switch target.
- Deleting an unrelated thread while a switch is pending.
- A stale new-thread result after a newer switch.
- A stale active-thread deletion after a newer switch.
- A thread selected after its deletion request began.
- A pending switch whose target is deleted before loading finishes.

- [x] **Step 4: Run focused tests and verify GREEN**

Run:

```powershell
cd D:\excel-ai-plugin-new-feature\desktop
npx vitest run src/store/chatStore.test.ts src/store/threadActions.test.ts
```

Expected: all tests pass, including the two new concurrency tests.

- [x] **Step 5: Run project verification**

Run:

```powershell
npm run lint
npm run typecheck
npm test
npm run build
```

Expected: all commands exit with code 0.

- [x] **Step 6: Check and commit the scoped diff**

Run:

```powershell
cd D:\excel-ai-plugin-new-feature
git diff --check
git status --short
git add -- desktop/src/store/chatStore.ts desktop/src/store/chatStore.test.ts docs/superpowers/plans/2026-07-10-thread-switch-race.md
git commit -m "fix: 防止旧会话加载覆盖最新切换"
```

Do not stage `.codex/`, `.workbuddy/screenshots/`, `desktop/data/`, or `desktop/vite-ui-refactor.log`.
