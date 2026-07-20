import { describe, expect, it, vi } from "vitest";
import { ChatController } from "../shared/agentChat";
import { ProviderStore } from "../shared/provider";
import { MockHostAdapter } from "./mockHost";

function sse(chunks: string[]): Response {
  const enc = new TextEncoder();
  let i = 0;
  return new Response(
    new ReadableStream<Uint8Array>({
      pull(c) {
        if (i >= chunks.length) {
          c.close();
          return;
        }
        c.enqueue(enc.encode(chunks[i++]));
      },
    }),
    { status: 200, headers: { "Content-Type": "text/event-stream" } },
  );
}

function d(obj: unknown): string {
  return `data: ${JSON.stringify(obj)}\n\n`;
}

const openaiTextStop = (text: string) => [
  d({
    id: "c",
    choices: [{ index: 0, delta: { content: text }, finish_reason: null }],
  }),
  d({
    id: "c",
    choices: [{ index: 0, delta: {}, finish_reason: "stop" }],
  }),
  "data: [DONE]\n\n",
];

function storeWith(apiKey = "sk-test-key") {
  const store = new ProviderStore();
  store.add({
    name: "openai",
    provider: "openai",
    apiKey,
    baseUrl: "https://api.openai.com/v1",
    model: "m",
    apiFormat: "openai",
  });
  return store;
}

function abortError(): Error {
  return Object.assign(new Error("aborted"), { name: "AbortError" });
}

/** Wait until predicate is true (microtask/macrotask polling for deterministic races). */
async function waitFor(
  predicate: () => boolean,
  label: string,
  attempts = 50,
): Promise<void> {
  for (let i = 0; i < attempts; i += 1) {
    if (predicate()) return;
    await new Promise<void>((r) => setTimeout(r, 0));
  }
  throw new Error(`waitFor timed out: ${label}`);
}

describe("ChatController stop and max_rounds", () => {
  it("A) stop during host probe: stopping → aborted, provider fetch never starts", async () => {
    let releaseStatus!: () => void;
    const statusGate = new Promise<void>((r) => {
      releaseStatus = r;
    });
    const host = new MockHostAdapter();
    const original = host.getStatus.bind(host);
    host.getStatus = async () => {
      await statusGate;
      return original();
    };

    const fetchImpl = vi.fn(async () => sse(openaiTextStop("should not run")));
    const controller = new ChatController({
      store: storeWith(),
      host,
      fetchImpl,
    });

    const turn = controller.send("probe-stop");
    await waitFor(
      () => controller.getState().status === "running",
      "running during host probe",
    );
    expect(fetchImpl).not.toHaveBeenCalled();

    // Concurrent send must be rejected while probe is in progress.
    const busy = await controller.send("second");
    expect(busy.turnStatus).toBe("busy");
    expect(fetchImpl).not.toHaveBeenCalled();

    controller.stop();
    expect(controller.getState().status).toBe("stopping");
    expect(fetchImpl).not.toHaveBeenCalled();

    releaseStatus();
    const result = await turn;
    expect(result.turnStatus).toBe("aborted");
    expect(result.error?.kind).toBe("aborted");
    expect(result.error?.message).toBe("aborted");
    expect(fetchImpl).not.toHaveBeenCalled();
    expect(controller.getState().status).toBe("idle");
    expect(controller.getState().lastTurnStatus).toBe("aborted");

    // Controller must accept a subsequent send (not stuck busy).
    fetchImpl.mockImplementation(async () => sse(openaiTextStop("ok")));
    const next = await controller.send("after");
    expect(next.turnStatus).toBe("completed");
    expect(controller.getState().status).toBe("idle");
  });

  it("B) stop after provider fetch started: signal aborts and turn ends idle", async () => {
    let seenSignal: AbortSignal | undefined;
    let fetchEntered!: () => void;
    const fetchStarted = new Promise<void>((r) => {
      fetchEntered = r;
    });

    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      const signal = init?.signal ?? undefined;
      seenSignal = signal;
      fetchEntered();
      if (signal?.aborted) throw abortError();
      await new Promise<void>((_resolve, reject) => {
        const onAbort = () => reject(abortError());
        // Must honor AbortSignal — do not resolve after stop.
        signal?.addEventListener("abort", onAbort, { once: true });
      });
      // Unreachable if abort is honored.
      return sse(openaiTextStop("too late"));
    });

    const controller = new ChatController({
      store: storeWith(),
      host: new MockHostAdapter(),
      fetchImpl,
    });

    const turn = controller.send("run");
    await fetchStarted;
    expect(fetchImpl).toHaveBeenCalledTimes(1);
    expect(seenSignal).toBeDefined();
    expect(seenSignal!.aborted).toBe(false);
    expect(controller.getState().status).toBe("running");

    controller.stop();
    expect(controller.getState().status).toBe("stopping");
    expect(seenSignal!.aborted).toBe(true);

    const result = await turn;
    expect(result.turnStatus).toBe("aborted");
    expect(result.error?.kind).toBe("aborted");
    expect(controller.getState().status).toBe("idle");
    expect(controller.getState().lastTurnStatus).toBe("aborted");
    // Fetch started once; abort prevented a successful completion body.
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it("max_rounds commits messages and reports max_rounds", async () => {
    const fetchImpl = vi.fn(async () =>
      sse([
        d({
          id: "c",
          choices: [
            {
              index: 0,
              delta: {
                tool_calls: [
                  {
                    index: 0,
                    id: "call_s",
                    type: "function",
                    function: { name: "host_status", arguments: "{}" },
                  },
                ],
              },
              finish_reason: null,
            },
          ],
        }),
        d({
          id: "c",
          choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }],
        }),
        "data: [DONE]\n\n",
      ]),
    );
    const controller = new ChatController({
      store: storeWith(),
      host: new MockHostAdapter(),
      fetchImpl,
      maxRounds: 1,
    });
    const result = await controller.send("status please");
    expect(result.turnStatus).toBe("max_rounds");
    expect(result.run?.rounds).toBe(1);
    expect(controller.getState().messages.length).toBeGreaterThan(0);
    expect(controller.getState().messages.some((m) => m.role === "user")).toBe(
      true,
    );
  });
});
